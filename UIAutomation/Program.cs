using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.MenuItems;

namespace ConsoleUIAutomation {
    class Program {
        static void Main(string[] args) {
            try {
                ProcessStartInfo pi = new ProcessStartInfo("notepad.exe");

                TestStack.White.Application app = TestStack.White.Application.AttachOrLaunch(pi);
                TestStack.White.UIItems.WindowItems.Window win = app.GetWindow("Untitled - Notepad");

                
                var item1 = win.Get<TestStack.White.UIItems.MenuItems.Menu>("File");
                var autoItem1 = item1.AutomationElement;
                
                item1.Focus();
                item1.Click();

                var menuOpen = win.Get<Menu>("Open...");
                
                if(menuOpen != null)
                Console.WriteLine(menuOpen.PrimaryIdentification);
                else
                    Console.WriteLine("NULL");
                
                //menuOpen.Focus();
                menuOpen.Click();
                var modalSavel = win.ModalWindow("Open");
                Thread.Sleep(590);
                modalSavel.Close();

                var sc = SearchCriteria.ByControlType(ControlType.Document);
                var doc = win.Get<TestStack.White.UIItems.TextBox>(sc);
                doc.Text = $"Some text {DateTime.Now.ToString()}";

                Thread.Sleep(590);
                win.Keyboard.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.CONTROL);
                win.Keyboard.Enter("s");
                win.Keyboard.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.CONTROL);
                var modalSaveAs = win.ModalWindow("Save As");

                var scFileName = SearchCriteria.ByControlType(ControlType.Edit).AndByClassName("Edit");
                var txtName = modalSaveAs.Get<TestStack.White.UIItems.TextBox>(scFileName);
                txtName.Text = $"newfile_{DateTime.UtcNow.ToString("yyyyMMdd_HHmmss_fff")}";
                var buttonSave = modalSaveAs.Get<TestStack.White.UIItems.Button>("Save");
                Thread.Sleep(590);
                buttonSave.Click();
                Console.WriteLine(item1.Id);
            } catch (Exception ex) {
                Console.Error.WriteLine(ex);
            }
        }
    }
}
