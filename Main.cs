using Ari.Importer.Properties;
using CsvHelper;
using Syncfusion.GridExcelConverter;
using Syncfusion.GridHelperClasses;
using Syncfusion.GroupingGridExcelConverter;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Grid.Grouping;
using Syncfusion.Windows.Forms.Tools;
using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;

namespace Ari.Importer
{
    public partial class Main : Form
    {
        private int childCount = 1;

        public Main()
        {
            InitializeComponent();
            this.Text = BL.Global.AppName;
        }

        private DataTable OpenCSV(string FilePathName)
        {
            var dt = new DataTable();

            using (var reader = new StreamReader(FilePathName))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                using (var dr = new CsvDataReader(csv))
                {
                    try
                    {
                        dt.Clear();
                        dt.Load(dr);
                        return dt;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                }
            }
        }

        private bool CreateGrid(DataTable dt)
        {
            GridGroupingControl dg = MainTabControl.SelectedTab.Controls.Find("dgFile" + MainTabControl.SelectedTab.Tag, true).FirstOrDefault() as GridGroupingControl;

            try
            {
                if (dt == null)
                {
                    return false;
                }
                else
                {
                    if (dt.Rows.Count > 0)
                    {
                        dg.DataSource = null;
                        dg.DataSource = dt;

                        string[] columnNames = dt.Columns.Cast<DataColumn>()
                                             .Select(x => x.ColumnName)
                                             .ToArray();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            dg.TopLevelGroupOptions.ShowFilterBar = true;
            foreach (GridColumnDescriptor column in dg.TableDescriptor.Columns)
            {
                column.AllowFilter = true;
            }

            GridExcelFilter filter = new GridExcelFilter
            {
                AllowResize = true,
                AllowSearch = true,
                EnableDateFilter = true,
                EnableNumberFilter = true,
                EnableStackedColumnFilterIcon = true
            };

            dg.TableControl.DpiAware = true;
            dg.TableOptions.AllowDragColumns = true;
            dg.TableOptions.AllowDropDownCell = true;
            dg.TableOptions.AllowMultiColumnSort = true;
            dg.TableOptions.AllowSelection = Syncfusion.Windows.Forms.Grid.GridSelectionFlags.Any;
            dg.TableOptions.AllowSortColumns = true;
            dg.TableOptions.ListBoxSelectionMode = SelectionMode.MultiExtended;
            dg.TopLevelGroupOptions.ShowAddNewRecordBeforeDetails = false;
            dg.TopLevelGroupOptions.ShowCaption = false;
            dg.GridVisualStyles = GridVisualStyles.Office2016Colorful;
            dg.NestedTableGroupOptions.ShowAddNewRecordBeforeDetails = false;
            dg.NestedTableGroupOptions.ShowCaption = false;
            dg.OptimizeFilterPerformance = true;
            dg.ShowNavigationBar = true;
            dg.ShowNavigationBarToolTips = true;


            try
            {
                filter.WireGrid(dg);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void SetupNewTabPage()
        {
            string FilePath = BL.Global.DefaultFolder;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                AddExtension = true,
                InitialDirectory = FilePath,
                Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true,
            };

            string File;
            string FileExt;
            string FileName;
            string FileNameAndExt;
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                File = openFileDialog.FileName;
                FileExt = Path.GetExtension(File);
                FileName = Path.GetFileNameWithoutExtension(File);
                FileNameAndExt = Path.GetFileName(File);
                BL.Manager.UpdateDefaultFolder(Path.GetDirectoryName(File));
            }
            else
            {
                return;
            }

            string ChildNumber = Convert.ToString(childCount++);
            string ChildName = "New " + ChildNumber;

            GridGroupingControl dg = new GridGroupingControl
            {
                Tag = ChildNumber,
                Name = "dgFile" + ChildNumber,
                HierarchicalGroupDropArea = true,
                ShowGroupDropArea = true,
                Dock = DockStyle.Fill,
            };

            TabPageAdv NewTabPage = new TabPageAdv
            {
                Tag = ChildNumber,
                Name = "tp" + ChildNumber,
                Text = ChildName,
                ToolTipText = FileName,
                AutoScroll = true,
            };

            NewTabPage.Controls.Add(dg);
            MainTabControl.TabPages.Add(NewTabPage);
            MainTabControl.SelectedTab = NewTabPage;

            switch (FileExt.ToLower())
            {
                case ".csv":
                    if (!CreateGrid(OpenCSV(File)))
                        CloseTab(false);
                    break;
                case ".txt":
                    if (!CreateGrid(OpenCSV(File)))
                        CloseTab(false);
                    break;
                case ".xls":
                    break;
                case ".xlsx":
                    break;
                default:
                    MessageBox.Show("This file type cannot be opened and parsed in this program. " + FileNameAndExt, "Error [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CloseTab(false);
                    break;
            }

            this.Text = BL.Global.AppName + " [" + ChildName + "]";
        }

        private void GroupAdded(object sender, EventArgs e)
        {
            //GridGroupingControl dg = MainTabControl.SelectedTab.Controls.Find("dgFile" + MainTabControl.SelectedTab.Tag, true).FirstOrDefault() as GridGroupingControl;
            //dg.Table.ExpandAllGroups();
        }

        private void CloseTab(bool CloseAllTabs)
        {
            if (MainTabControl.TabPages.Count > 0)
            {
                if (CloseAllTabs)
                {
                    foreach (TabPageAdv tab in MainTabControl.TabPages)
                    {
                        MainTabControl.TabPages.Remove(tab);
                    }

                    this.Text = BL.Global.AppName;
                }
                else
                {
                    MainTabControl.TabPages.Remove(MainTabControl.SelectedTab);

                    if (MainTabControl.TabPages.Count > 0)
                        this.Text = BL.Global.AppName + " [" + MainTabControl.SelectedTab.Text + "]";
                    else
                        this.Text = BL.Global.AppName;
                }
            }
        }

        private void SaveGrid(String FileName)
        {
            GridGroupingControl dg;

            try
            {
                dg = MainTabControl.SelectedTab.Controls.Find("dgFile" + MainTabControl.SelectedTab.Tag, true).FirstOrDefault() as GridGroupingControl;
            }
            catch
            {
                return;
            }

            if (FileName == string.Empty)
            {
                string FilePath = BL.Global.DefaultFolder;

                FileDialog saveFileDlg = new SaveFileDialog
                {
                    AddExtension = true,
                    InitialDirectory = FilePath,
                    Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,
                    //Title = "Save Grid Settings As"
                };

                if (saveFileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = saveFileDlg.FileName;
                    BL.Manager.UpdateDefaultFolder(Path.GetDirectoryName(FileName));
                }
                else
                {
                    return;
                }
            }

            // Create xml writer for adding the info to the xml file.
            XmlTextWriter xmlWriter = new XmlTextWriter(FileName, System.Text.Encoding.UTF8)
            {
                Formatting = Formatting.Indented
            };

            // Write Grid schema to the xml file.
            dg.WriteXmlSchema(xmlWriter);
            xmlWriter.Close();

            MessageBox.Show("Settings have been saved to " + FileName + ".", "Settings Saved [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ApplyGridSettings()
        {
            GridGroupingControl dg;

            try
            {
                dg = MainTabControl.SelectedTab.Controls.Find("dgFile" + MainTabControl.SelectedTab.Tag, true).FirstOrDefault() as GridGroupingControl;
            }
            catch
            {
                return;
            }

            string FilePath = BL.Global.DefaultFolder;

            FileDialog openFileDlg = new OpenFileDialog
            {
                AddExtension = true,
                InitialDirectory = FilePath,
                Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true,
                //Title = "Open Grid Settings"
            };

            if (openFileDlg.ShowDialog() == DialogResult.OK)
            {
                BL.Manager.UpdateDefaultFolder(Path.GetDirectoryName(openFileDlg.FileName));

                XmlReader xmlReader = new XmlTextReader(openFileDlg.FileName);
                dg.ApplyXmlSchema(xmlReader);
                dg.GridVisualStyles = GridVisualStyles.Office2016Colorful;
                xmlReader.Close();

                MessageBox.Show("Settings from " + openFileDlg.FileName + " have been applied.", "Settings Applied [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                return;
            }
        }

        private void PrintPreview()
        {
            string FilePath = BL.Global.DefaultFolder;

            SaveFileDialog saveFileDlg = new SaveFileDialog
            {
                AddExtension = true,
                InitialDirectory = FilePath,
                Filter = "PDF|*.pdf|Excel|*.xls",
                FilterIndex = 1,
                RestoreDirectory = true,
                //Title = "Save Print Output As"
            };
            saveFileDlg.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (saveFileDlg.FileName != "")
            {
                try
                {
                    GridGroupingControl dg = MainTabControl.SelectedTab.Controls.Find("dgFile" + MainTabControl.SelectedTab.Tag, true).FirstOrDefault() as GridGroupingControl;

                    // File type selected in the dialog box.
                    switch (saveFileDlg.FilterIndex)
                    {
                        case 1:
                            // Export the contents of the Grid to Pdf
                            GridPDFConverter pdfConvertor = new GridPDFConverter();
                            pdfConvertor.ExportToPdf(saveFileDlg.FileName, dg.TableControl);
                            break;

                        case 2:
                            // Export the contents of the Grid to Excel
                            GroupingGridExcelConverterControl converter = new GroupingGridExcelConverterControl();
                            converter.GroupingGridToExcel(dg, saveFileDlg.FileName, ConverterOptions.Visible);
                            break;
                    }

                    MessageBox.Show("File has been saved to " + saveFileDlg.FileName + ".", "Print Output [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error [" + BL.Global.AppName + "]", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void MainTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (MainTabControl.TabPages.Count > 0)
            {
                this.Text = BL.Global.AppName + " [" + MainTabControl.SelectedTab.Text + "]";
            }
            else
            {
                this.Text = BL.Global.AppName;
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void NewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetupNewTabPage();
        }

        private void NewToolStripButton_Click(object sender, EventArgs e)
        {
            SetupNewTabPage();
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ApplyGridSettings();
        }

        private void OpenToolStripButton_Click(object sender, EventArgs e)
        {
            ApplyGridSettings();
        }

        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveGrid(string.Empty);
        }

        private void SaveToolStripButton_Click(object sender, EventArgs e)
        {
            SaveGrid(string.Empty);
        }

        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveGrid(string.Empty);
        }

        private void SaveAsToolStripButton_Click(object sender, EventArgs e)
        {
            SaveGrid(string.Empty);
        }

        private void CloseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseTab(false);
        }

        private void CloseToolStripButton_Click(object sender, EventArgs e)
        {
            CloseTab(false);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseTab(true);
        }

        private void CloseAllToolStripButton_Click(object sender, EventArgs e)
        {
            CloseTab(true);
        }

        private void PrintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintPreview();
        }

        private void PrintToolStripButton_Click(object sender, EventArgs e)
        {
            PrintPreview();
        }

        private void PrintPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This feature not implemented yet.", BL.Global.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void PrintPreviewToolStripButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This feature not implemented yet.", BL.Global.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ToolBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStrip.Visible = toolBarToolStripMenuItem.Checked;
            Settings.Default.ToolBar = toolBarToolStripMenuItem.Checked;
            Settings.Default.Save();
        }

        private void StatusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            statusStrip.Visible = statusBarToolStripMenuItem.Checked;
            Settings.Default.StatusBar = statusBarToolStripMenuItem.Checked;
            Settings.Default.Save();
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox Form = new AboutBox();
            Form.ShowDialog();
        }

        private void OptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This feature not implemented yet.", BL.Global.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Main_Load(object sender, EventArgs e)
        {
            // Set window location
            if (Settings.Default.WindowLocation != null)
            {
                this.Location = Settings.Default.WindowLocation;
            }

            // Set window size
            if (Settings.Default.WindowSize != null)
            {
                this.Size = Settings.Default.WindowSize;
            }

            // Set toolbar status
            bool? bToolBar = Settings.Default.ToolBar;
            if (bToolBar.HasValue)
            {
                toolStrip.Visible = Settings.Default.ToolBar;
                toolBarToolStripMenuItem.Checked = Settings.Default.ToolBar;
            }

            // Set statusbar status
            bool? bstatusbar = Settings.Default.StatusBar;
            if (bstatusbar.HasValue)
            {
                statusStrip.Visible = Settings.Default.StatusBar;
                statusBarToolStripMenuItem.Checked = Settings.Default.StatusBar;
            }

            //Make sure visible on screen
            if (BL.Manager.IsOnScreen(this) != true)
            {
                this.Location = new Point(25, 25);
                this.Size = new Size(870, 653);
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason.ToString() != "ApplicationExitCall")
            {
                if (DialogResult.Yes == MessageBox.Show("Are you sure you wish to exit?", "Confirm Exit [" + BL.Global.AppName + "]", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
                {
                    // Copy window location to app settings
                    Settings.Default.WindowLocation = this.Location;

                    // Copy window size to app settings
                    if (this.WindowState == FormWindowState.Normal)
                    {
                        Settings.Default.WindowSize = this.Size;
                    }
                    else
                    {
                        Settings.Default.WindowSize = this.RestoreBounds.Size;
                    }

                    // Save settings
                    Settings.Default.Save();

                    Application.Exit();
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }
    }
}
