using Ari.Importer.Properties;
using System.Drawing;
using System.Windows.Forms;

namespace Ari.Importer.BL
{
    class Manager
    {
        public static void UpdateDefaultFolder(string folder)
        {
            BL.Global.DefaultFolder = folder;
            Settings.Default.DefaultFolder = folder;
            Settings.Default.Save();
        }

        public static bool IsOnScreen(Form form)
        {
            Screen[] screens = Screen.AllScreens;
            foreach (Screen screen in screens)
            {
                Rectangle formRectangle = new Rectangle(form.Left, form.Top, form.Width, form.Height);
                if (screen.WorkingArea.Contains(formRectangle))
                {
                    return true;
                }
            }

            return false;
        }

    }
}
