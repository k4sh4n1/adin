using KD.SDK2;

using System.Windows.Forms;

using System.IO;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System;

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

                var xmlDocument = new XmlDocument();

                var sceneXmlElement = xmlDocument.CreateElement("Scene");
                xmlDocument.AppendChild(sceneXmlElement);

                sceneXmlElement.SetAttribute("FilenameWithPath", scene.FilenameWithPath);

                var customerXmlElement = xmlDocument.CreateElement("Customer");
                sceneXmlElement.AppendChild(customerXmlElement);

                customerXmlElement.SetAttribute("Company",
                                                scene.KeywordInfo["@Base.CustomerCompany()"]);
                customerXmlElement.SetAttribute("Name",
                                                scene.KeywordInfo["@Base.CustomerName()"]);
                customerXmlElement.SetAttribute("FirstName",
                                                scene.KeywordInfo["@Base.CustomerFirstName()"]);

                foreach (Scene.Object obj in scene.Objects)
                {
                    if (obj.CatalogFilename == "SYSTEM" || obj.CatalogFilename.StartsWith("@"))
                        continue;

                    var objXmlElement = xmlDocument.CreateElement("Object");
                    sceneXmlElement.AppendChild(objXmlElement);

                    objXmlElement.SetAttribute("Number", obj.Number);
                    objXmlElement.SetAttribute("Type", obj.Type.ToString());
                    objXmlElement.SetAttribute("Catalog", obj.CatalogFilename.ToString());
                    objXmlElement.SetAttribute("KeyReference", obj.KeyReference);
                    objXmlElement.SetAttribute("Handing", obj.Handing.ToString());

                    var nameXmlElement = xmlDocument.CreateElement("Name");
                    objXmlElement.AppendChild(nameXmlElement);

                    nameXmlElement.InnerText = obj.Name;

                    if (obj.SupplierComment != string.Empty)
                    {
                        var supplierCommentXmlElement = xmlDocument.CreateElement("SupplierComment");
                        objXmlElement.AppendChild(supplierCommentXmlElement);

                        supplierCommentXmlElement.InnerText = obj.FitterComment;
                    }

                    var objFinishesConfig = obj.GetFinishesConfig();
                    foreach (var finish in objFinishesConfig.Finishes)
                    {
                        var finishXmlElement = xmlDocument.CreateElement("Finish");
                        objXmlElement.AppendChild(finishXmlElement);

                        finishXmlElement.SetAttribute("Type", finish.Type.Name);
                        finishXmlElement.SetAttribute("Code", finish.Code);
                        finishXmlElement.SetAttribute("Name", finish.Name);
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XML-File | *.xml";
                saveFileDialog.Title = "Save As XML File";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    xmlDocument.Save(saveFileDialog.FileName);
                }
                else
                {
                    MessageBox.Show("Please choose a proper XML file name");
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
