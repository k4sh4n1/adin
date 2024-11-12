using KD.SDK2;

using System.Windows.Forms;

using System.IO;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System;
using System.Text;

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
            try
            {
                Scene scene = _appli.Scene;

                StringBuilder csvContent = new StringBuilder();

                // Add header row
                csvContent.AppendLine("Number,Type,Catalog,KeyReference,Handing,Name,SupplierComment,FinishType,FinishCode,FinishName");

                foreach (Scene.Object obj in scene.Objects)
                {
                    string supplierComment = !string.IsNullOrEmpty(obj.SupplierComment) ? obj.SupplierComment : string.Empty;

                    var objFinishesConfig = obj.GetFinishesConfig();
                    if (objFinishesConfig.Finishes.Count() == 0)
                    {
                        // Add a single row if no finishes exist
                        csvContent.AppendLine($"{obj.Number},{obj.Type},{obj.CatalogFilename},{obj.KeyReference},{obj.Handing},{obj.Name},{supplierComment},,,");
                    }
                    else
                    {
                        // Add a row for each finish
                        foreach (var finish in objFinishesConfig.Finishes)
                        {
                            csvContent.AppendLine($"{obj.Number},{obj.Type},{obj.CatalogFilename},{obj.KeyReference},{obj.Handing},{obj.Name},{supplierComment},{finish.Type.Name},{finish.Code},{finish.Name}");
                        }
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV-File | *.csv";
                saveFileDialog.Title = "Save As CSV File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog.FileName, csvContent.ToString());
                }
                else
                {
                    MessageBox.Show("Please choose a proper CSV file name");
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                string.Join(Environment.NewLine,
                    string.Format("Failed to run plugin '{0}':", "Adin"),
                    exception.Message));
                return false;
            }

            return true;
        }
    }
}
