using KD.SDK2;

using System.Windows.Forms;

using System.IO;
using System.Reflection;

namespace Adin
{
    public class Plugin
    {
        Appli _appli;
        Appli.MenuItem _myMenuItem;

        public Plugin()
        {
            _appli = new Appli();
        }

        public bool OnPluginLoad(int unused)
        {
            var menuInfo = new Appli.MenuItemInsertionInfo();

            menuInfo.Text = "Adin";
            menuInfo.DllFilenameWithoutPath = Path.GetFileName(Assembly.GetExecutingAssembly().Location);
            menuInfo.ClassName = nameof(Plugin);
            menuInfo.MethodName = nameof(OnMyMenuItem);

            _myMenuItem = _appli.InsertMenuItem(menuInfo, Appli.MenuItem.StandardId.File_ExportBmp);

            return true;
        }

        public bool OnPluginUnload(int unused)
        {
            _appli.RemoveMenuItem(_myMenuItem);

            return true;
        }

        public bool OnMyMenuItem(int unused)
        {
            MessageBox.Show("Enjoy.");

            return true;
        }
    }
}
