using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System.Windows.Media;

namespace TestRibbon
{
    public class MyRibbon : IExtensionApplication
    {
        
        public void Initialize()
        {
            //Autodesk.Windows.ComponentManager.ItemInitialized += new EventHandler(ComponentManager_ItemInitialized);
            BuildRibbonTab();
        }

        public void Terminate()
        {
            
        }

        void ComponentManager_ItemInitialized(object sender, Autodesk.Windows.RibbonItemEventArgs e)
        {
            // Проверяем, что лента загружена
            if (Autodesk.Windows.ComponentManager.Ribbon != null)
            {
                // Строим нашу вкладку
                BuildRibbonTab();
                //и раз уж лента запустилась, то отключаем обработчик событий
                Autodesk.Windows.ComponentManager.ItemInitialized -= new EventHandler<RibbonItemEventArgs>(ComponentManager_ItemInitialized);
            }
        }

        private void BuildRibbonTab()
        {
            // Если лента еще не загружена
            if (!RibbonisLoaded())
            {
                // Строим вкладку
                CreateRibbonTab();
                // Подключаем обработчик событий изменения системных переменных
                AcadApp.SystemVariableChanged += new SystemVariableChangedEventHandler(AcadApp_SystemVariableChanged);
            }
        }

        private void AcadApp_SystemVariableChanged(object sender, SystemVariableChangedEventArgs e)
        {
            if (e.Name.Equals("WSCURRENT")) BuildRibbonTab();
        }

        bool RibbonisLoaded()
        {
            bool _loaded = false;
            RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
            // Делаем итерацию по вкладкам ленты
            foreach (RibbonTab tab in ribCntrl.Tabs)
            {
                // И если у вкладки совпадает идентификатор и заголовок, то значит вкладка загружена
                if (tab.Id.Equals("RibbonExample_ID") & tab.Title.Equals("RibbonExample"))
                { _loaded = true; break; }
                else _loaded = false;
            }
            return _loaded;
        }
        private void CreateRibbonTab()
        {
            try
            {
                // Получаем доступ к ленте
                RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
                // добавляем свою вкладку
                RibbonTab ribTab = new RibbonTab();
                ribTab.Title = "RibbonExample"; // Заголовок вкладки
                ribTab.Id = "RibbonExample_ID"; // Идентификатор вкладки
                ribCntrl.Tabs.Add(ribTab); // Добавляем вкладку в ленту
                // добавляем содержимое в свою вкладку (одну панель)
                AddExampleContent(ribTab);
                // Делаем вкладку активной (не желательно, ибо неудобно)
                //ribTab.IsActive = true;
                // Обновляем ленту (если делаете вкладку активной, то необязательно)
                ribCntrl.UpdateLayout();
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.
                  DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
            }
        }

        private void AddExampleContent(RibbonTab ribTab)
        {
            try
            {
                // создаем panel source
                RibbonPanelSource ribSourcePanel = new RibbonPanelSource();
                ribSourcePanel.Title = "RibbonExample";
                // теперь саму панель
                RibbonPanel ribPanel = new RibbonPanel();
                ribPanel.Source = ribSourcePanel;
                ribTab.Panels.Add(ribPanel);
                // создаем пустую tooltip (всплывающая подсказка)
                RibbonToolTip tt;
                // создаем split button
                RibbonSplitButton risSplitBtn = new RibbonSplitButton();
                /* Для RibbonSplitButton ОБЯЗАТЕЛЬНО надо указать
                 * свойство Text, а иначе при поиске команд в автокаде
                 * будет вылетать ошибка.
                 */
                risSplitBtn.Text = "RibbonSplitButton";
                // Ориентация кнопки
                risSplitBtn.Orientation = System.Windows.Controls.Orientation.Vertical;
                // Размер кнопки
                risSplitBtn.Size = RibbonItemSize.Large;
                // Показывать изображение
                risSplitBtn.ShowImage = true;
                // Показывать текст
                risSplitBtn.ShowText = true;
                // Стиль кнопки
                risSplitBtn.ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton;
                risSplitBtn.ResizeStyle = RibbonItemResizeStyles.NoResize;
                risSplitBtn.ListStyle = RibbonSplitButtonListStyle.List;
                /* Далее создаем две кнопки и добавляем их
                 * не в панель, а в RibbonSplitButton
                 */
                #region Кнопка-пример №1
                // Создаем новый экземпляр подсказки
                tt = new RibbonToolTip();
                // Отключаем вызов справки (в данном примере её нету)
                tt.IsHelpEnabled = false;
                // Создаем кнопку
                RibbonButton ribBtn = new RibbonButton();
                /* В свойство CommandParameter (параметры команды)
                 * и в свойство Command (отображает команду) подсказки
                 * пишем вызываемую команду
                 */
                ribBtn.CommandParameter = tt.Command = "_Line";
                // Имя кнопки
                ribBtn.Name = "ExampleButton1";
                // Заголовок кнопки и подсказки
                ribBtn.Text = tt.Title = "Кнопка-пример №1";
                // Создаем новый (собственный) обработчик команд (см.ниже)
                ribBtn.CommandHandler = new RibbonCommandHandler();
                // Ориентация кнопки
                ribBtn.Orientation = System.Windows.Controls.Orientation.Horizontal;
                // Размер кнопки
                ribBtn.Size = RibbonItemSize.Large;
                /* Т.к. используем размер кнопки Large, то добавляем
                 * большое изображение с помощью специальной функции (см.ниже)
                 */
                ribBtn.LargeImage = LoadImage("mind_map_32px");
                // Показывать картинку
                ribBtn.ShowImage = true;
                // Показывать текст
                ribBtn.ShowText = true;
                // Заполняем содержимое всплывающей подсказки
                tt.Content = "Я кнопочка №1. Нажми меня и я нарисую отрезок";
                // Подключаем подсказку к кнопке
                ribBtn.ToolTip = tt;
                // Добавляем кнопку в RibbonSplitButton
                risSplitBtn.Items.Add(ribBtn);
                #endregion
                // Делаем текущей первую кнопку
                risSplitBtn.Current = ribBtn;
                // Далее создаем вторую кнопку по аналогии с первой
                #region Кнопка-пример №2
                tt = new RibbonToolTip();
                tt.IsHelpEnabled = false;
                ribBtn = new RibbonButton();
                ribBtn.CommandParameter = tt.Command = "_Pline";
                ribBtn.Name = "ExampleButton2";
                ribBtn.Text = tt.Title = "Кнопка-пример №2";
                ribBtn.CommandHandler = new RibbonCommandHandler();
                ribBtn.Orientation = System.Windows.Controls.Orientation.Horizontal;
                ribBtn.Size = RibbonItemSize.Large;
                ribBtn.LargeImage = LoadImage("mind_map_32px");
                ribBtn.ShowImage = true;
                ribBtn.ShowText = true;
                tt.Content = "Я кнопочка №2. Нажми меня и я нарисую полилинию";
                ribBtn.ToolTip = tt;
                risSplitBtn.Items.Add(ribBtn);
                #endregion
                // Добавляем RibbonSplitButton в нашу панель
                ribSourcePanel.Items.Add(risSplitBtn);
                // Создаем новую строку
                RibbonRowPanel ribRowPanel = new RibbonRowPanel();
                // Создаем третью кнопку по аналогии с предыдущими.
                // Отличие только в размере кнопки (и картинки)
                #region Кнопка-пример №3
                tt = new RibbonToolTip();
                tt.IsHelpEnabled = false;
                ribBtn = new RibbonButton();
                ribBtn.CommandParameter = tt.Command = "_Circle";
                ribBtn.Name = "ExampleButton3";
                ribBtn.Text = tt.Title = "Кнопка-пример №3";
                ribBtn.CommandHandler = new RibbonCommandHandler();
                ribBtn.Orientation = System.Windows.Controls.Orientation.Vertical;
                ribBtn.Size = RibbonItemSize.Standard;
                ribBtn.Image = LoadImage("mind_map_16px");
                ribBtn.ShowImage = true;
                ribBtn.ShowText = false;
                tt.Content = "Я кнопочка №3. Нажми меня и я нарисую кружочек";
                ribBtn.ToolTip = tt;
                ribRowPanel.Items.Add(ribBtn);
                #endregion
                // Добавляем строку в нашу панель
                ribSourcePanel.Items.Add(ribRowPanel);
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.
                  DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
            }
        }

        private ImageSource LoadImage(string resourceName)
        {
            ImageSource imageSource = new System.Windows.Media.Imaging.BitmapImage(new Uri("pack://application:,,,/TestRibbon;component/" + resourceName + ".png"));
            return imageSource;
        }


        /* Собственный обраотчик команд
         * Это один из вариантов вызова команды по нажатию кнопки
         */
        class RibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object parameter)
            {
                return true;
            }
            public event EventHandler CanExecuteChanged;
            public void Execute(object parameter)
            {
                TestWindow window = new TestWindow();
                window.Show();
                //Document doc = AcadApp.DocumentManager.MdiActiveDocument;
                //if (parameter is RibbonButton)
                //{
                //    // Просто берем команду, записанную в CommandParameter кнопки
                //    // и выпоняем её используя функцию SendStringToExecute
                //    RibbonButton button = parameter as RibbonButton;
                //    AcadApp.DocumentManager.MdiActiveDocument.SendStringToExecute(button.CommandParameter + " ", true, false, true);
                //}
            }
        }
    }
}
