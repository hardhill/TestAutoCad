using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using HelloCad.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelloRibbon
{
    public class HelloAutocad:IExtensionApplication
    {

        private HelloWindow helloWindow;
        public HelloAutocad()
        {
            helloWindow = new HelloWindow();
        }

        [CommandMethod("TestHello")]
        public void TestHello()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Привет из Autocad плагина");
            helloWindow.Show();
        }

        public void Initialize()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Инициализация плагина.." + Environment.NewLine);
        }

        public void Terminate()
        {
            
        }
    }
}
