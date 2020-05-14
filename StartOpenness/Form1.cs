using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Siemens.Engineering.SW.Blocks;
using System.Diagnostics;
using System.Xml.Linq;
using System.Linq;

namespace StartOpenness
{
    public partial class Form1 : Form
    {


        private static TiaPortalProcess _tiaProcess;
        private static Project _tiaProject;
        private static PlcSoftware _controller;


        private static OpenFileDialog _OFD_Import = new OpenFileDialog();
        private static FolderBrowserDialog _FBD_Export = new FolderBrowserDialog();

        private static FolderBrowserDialog _FBD_Path = new FolderBrowserDialog();

        private int rowIndex = 0; 


        struct Mechanism
        {
            public string name;
            public string type;
        };


        public TiaPortal MyTiaPortal
        {
            get; set;
        }
        public Project MyProject
        {
            get; set;
        }


        public Form1()
        {
            InitializeComponent();
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
        }



        private void Form1_Load_1(object sender, EventArgs e)
        {
            _FBD_Export.Description = "Select a folder to ";
            _FBD_Export.SelectedPath = Properties.Settings.Default.Export_Path; // persisted
            textExport.Text = _FBD_Export.SelectedPath;

            _FBD_Path.Description = "Select a folder with your own FB`s";
            _FBD_Path.SelectedPath = Properties.Settings.Default.FB_Repository;
            update_Datagridview();
        }

        private static Assembly MyResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf(',');
            if (index == -1)
            {
                return null;
            }
            string name = args.Name.Substring(0, index);

            RegistryKey filePathReg = Registry.LocalMachine.OpenSubKey(
                "SOFTWARE\\Siemens\\Automation\\Openness\\14.0\\PublicAPI\\14.0.1.0");

            object oRegKeyValue = filePathReg?.GetValue(name);
            if (oRegKeyValue != null)
            {
                string filePath = oRegKeyValue.ToString();

                string path = filePath;
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return Assembly.LoadFrom(fullPath);
                }
            }

            return null;
        }


        private void StartTIA(object sender, EventArgs e)
        {
            if (rdb_WithoutUI.Checked == true)
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                txt_Status.Text = "TIA Portal started without user interface";
                _tiaProcess = TiaPortal.GetProcesses()[0];
            }
            else
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
                txt_Status.Text = "TIA Portal started with user interface";
            }

            btn_SearchProject.Enabled = true;
            btn_Dispose.Enabled = true;
            btn_Start.Enabled = false;

        }

        private void DisposeTIA(object sender, EventArgs e)
        {
            MyTiaPortal.Dispose();
            txt_Status.Text = "TIA Portal disposed";

            btn_Start.Enabled = true;
            btn_Dispose.Enabled = false;
            btn_CloseProject.Enabled = false;
            btn_SearchProject.Enabled = false;
            btn_CompileHW.Enabled = false;
            btn_Save.Enabled = false;


        }

        private void SearchProject(object sender, EventArgs e)
        {

            OpenFileDialog fileSearch = new OpenFileDialog();

            fileSearch.Filter = "*.ap14|*.ap14";
            fileSearch.RestoreDirectory = true;
            fileSearch.ShowDialog();

            string ProjectPath = fileSearch.FileName.ToString();

            if (string.IsNullOrEmpty(ProjectPath) == false)
            {
                OpenProject(ProjectPath);
            }
        }

        private void OpenProject(string ProjectPath)
        {
            try
            {
                MyProject = MyTiaPortal.Projects.Open(new FileInfo(ProjectPath));
                txt_Status.Text = "Project " + ProjectPath + " opened";

            }
            catch (Exception ex)
            {
                txt_Status.Text = "Error while opening project" + ex.Message;
            }

            btn_CompileHW.Enabled = true;
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
            btn_AddHW.Enabled = true;
        }

        private void SaveProject(object sender, EventArgs e)
        {
            MyProject.Save();
            txt_Status.Text = "Project saved";
        }


        private void CloseProject(object sender, EventArgs e)
        {
            MyProject.Close();

            txt_Status.Text = "Project closed";

            btn_SearchProject.Enabled = true;
            btn_CloseProject.Enabled = false;
            btn_Save.Enabled = false;
            btn_CompileHW.Enabled = false;


        }

        private void Compile(object sender, EventArgs e)
        {
            btn_CompileHW.Enabled = false;

            string devname = txt_Device.Text;
            bool found = false;

            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    if (deviceItem.Name == devname || device.Name == devname)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            if (softwareContainer.Software is PlcSoftware)
                            {
                                PlcSoftware controllerTarget = softwareContainer.Software as PlcSoftware;
                                if (controllerTarget != null)
                                {
                                    found = true;
                                    ICompilable compiler = controllerTarget.GetService<ICompilable>();

                                    CompilerResult result = compiler.Compile();
                                    txt_Status.Text = "Compiling of " + controllerTarget.Name + ": State: " + result.State + " / Warning Count: " + result.WarningCount + " / Error Count: " + result.ErrorCount;

                                }
                            }
                            if (softwareContainer.Software is HmiTarget)
                            {
                                HmiTarget hmitarget = softwareContainer.Software as HmiTarget;
                                if (hmitarget != null)
                                {
                                    found = true;
                                    ICompilable compiler = hmitarget.GetService<ICompilable>();
                                    CompilerResult result = compiler.Compile();
                                    txt_Status.Text = "Compiling of " + hmitarget.Name + ": State: " + result.State + " / Warning Count: " + result.WarningCount + " / Error Count: " + result.ErrorCount;
                                }

                            }
                        }
                    }
                }
            }
            if (found == false)
            {
                txt_Status.Text = "Found no device with name " + txt_Device.Text;
            }

            btn_CompileHW.Enabled = true;
        }

        private void btn_AddHW_Click(object sender, EventArgs e)
        {
            btn_AddHW.Enabled = false;
            string MLFB = "OrderNumber:" + txt_OrderNo.Text + "/" + txt_Version.Text;

            string name = txt_AddDevice.Text;
            string devname = "station" + txt_AddDevice.Text;
            bool found = false;
            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    if (deviceItem.Name == devname || device.Name == devname)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            if (softwareContainer.Software is PlcSoftware)
                            {
                                PlcSoftware controllerTarget = softwareContainer.Software as PlcSoftware;
                                if (controllerTarget != null)
                                {
                                    found = true;

                                }
                            }
                            if (softwareContainer.Software is HmiTarget)
                            {
                                HmiTarget hmitarget = softwareContainer.Software as HmiTarget;
                                if (hmitarget != null)
                                {
                                    found = true;

                                }

                            }
                        }
                    }
                }
            }
            if (found == true)
            {
                txt_Status.Text = "Device " + txt_Device.Text + " already exists";
            }
            else
            {
                Device deviceName = MyProject.Devices.CreateWithItem(MLFB, name, devname);

                txt_Status.Text = "Add Device Name: " + name + " with Order Number: " + txt_OrderNo.Text + " and Firmware Version: " + txt_Version.Text;
            }

            btn_AddHW.Enabled = true;

        }

        private void btn_ConnectTIA(object sender, EventArgs e)
        {
            btn_Connect.Enabled = false;
            IList<TiaPortalProcess> processes = TiaPortal.GetProcesses();
            switch (processes.Count)
            {
                case 1:
                    _tiaProcess = processes[0];
                    MyTiaPortal = _tiaProcess.Attach();
                    if (MyTiaPortal.GetCurrentProcess().Mode == TiaPortalMode.WithUserInterface)
                    {
                        rdb_WithUI.Checked = true;
                    }
                    else
                    {
                        rdb_WithoutUI.Checked = true;
                    }


                    if (MyTiaPortal.Projects.Count <= 0)
                    {
                        txt_Status.Text = "No TIA Portal Project was found!";
                        btn_Connect.Enabled = true;
                        return;
                    }
                    MyProject = MyTiaPortal.Projects[0];
                    break;
                case 0:
                    txt_Status.Text = "No running instance of TIA Portal was found!";
                    btn_Connect.Enabled = true;
                    return;
                default:
                    txt_Status.Text = "More than one running instance of TIA Portal was found!";
                    btn_Connect.Enabled = true;
                    return;
            }
            txt_Status.Text = _tiaProcess.ProjectPath.ToString();
            btn_Start.Enabled = false;
            btn_Connect.Enabled = true;
            btn_Dispose.Enabled = true;
            btn_CompileHW.Enabled = true;
            btn_CloseProject.Enabled = true;
            btn_SearchProject.Enabled = false;
            btn_Save.Enabled = true;
            btn_AddHW.Enabled = true;
            btn_Import.Enabled = true;
            btn_Export.Enabled = true;

        }

        private void Btn_ExportBlock(object sender, EventArgs e)
        {
            FindPlc();
            ExportBlocks(_controller);
            txt_Status.Text = "Block Exported";
        }

        private void Btn_ImportBlock(object sender, EventArgs e)
        {
            FindPlc();
            ImportBlocks(_controller, _OFD_Import.FileName);
            txt_Status.Text = "Block Imported";
        }

        public void FindPlc()
        {
            _controller = null;
            PlcSoftware firstPlc = null;
            var number = 0;
            var devices = MyProject.Devices;
            foreach (var device in devices)
            {
                Console.WriteLine(@"Device:	" + device.Name);

                if (device.Name != "SINAMICS G120")
                {
                    foreach (var item in device.DeviceItems)
                    {
                        var deviceItem = (DeviceItem)item;
                        //var plcSoftware = deviceItem.GetService<ISoftwareContainer>(); TIA Portal V14

                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();

                        if (softwareContainer != null)

                        {

                            Software softwareBase = softwareContainer.Software;

                            PlcSoftware plcSoftware = softwareBase as PlcSoftware;


                            if (plcSoftware != null)
                            {

                                Console.WriteLine(@"Software gefunden in: " + deviceItem.Name);
                                _controller = plcSoftware;
                                break;

                            }
                        }
                    }
                }
                if (_controller != null)
                {
                    break;
                }
            }
            if (_controller == null)
            {
                if (firstPlc != null)
                {
                    if (number > 1)
                        throw new SystemException(
                            "To much ControllerTarget in project");
                    _controller = firstPlc;
                }
                else
                {
                    throw new SystemException("No ControllerTarget in project.");
                }
            }
        }

        private  void ExportBlocks(PlcSoftware plcSoftware)
        {

            PlcBlock _plcblock = plcSoftware.BlockGroup.Blocks.Find(txt_BlockName.Text);
            //PlcBlock _plcblock = plcSoftware.BlockGroup.Blocks.Find("Empty_FC");
            _plcblock.Export(new FileInfo(string.Format(_FBD_Export.SelectedPath + @"\" + "{0}.xml", _plcblock.Name)), ExportOptions.WithDefaults);
        }



        private  void ImportBlocks(PlcSoftware plcSoftware, string path)
        {
            //string path = _OFD_Import.FileName;
            PlcBlockGroup blockGroup = plcSoftware.BlockGroup;
            IList<PlcBlock> blocks = blockGroup.Blocks.Import(new FileInfo(path), ImportOptions.Override);
        }

       

        private void Btn_Browse_Click(object sender, EventArgs e)
        {
            if (_FBD_Export.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.Export_Path = _FBD_Export.SelectedPath;
                Properties.Settings.Default.Save();
                if (!string.IsNullOrEmpty(Properties.Settings.Default.Export_Path))
                    Properties.Settings.Default.Export_Path = _FBD_Export.SelectedPath;
            }
            textExport.Text = _FBD_Export.SelectedPath;


        }

        private void Btn_OpenImport_Click(object sender, EventArgs e)
        {
            //OFD_Import.InitialDirectory = "C:\\";
            _OFD_Import.Filter = "xml files (*.xml)|*.xml";
            _OFD_Import.RestoreDirectory = true;

            if (_OFD_Import.ShowDialog() == DialogResult.OK)
                textImport.Text = _OFD_Import.FileName;
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
               if(_FBD_Path.ShowDialog() == DialogResult.OK)
               {
                Properties.Settings.Default.FB_Repository = _FBD_Path.SelectedPath;
                Properties.Settings.Default.Save();

                    if(!string.IsNullOrEmpty(Properties.Settings.Default.FB_Repository))
                    Properties.Settings.Default.FB_Repository = _FBD_Path.SelectedPath;
               
                }
            update_Datagridview();
            


        }

        private void update_Datagridview()
        {
            DirectoryInfo dInfo = new DirectoryInfo(_FBD_Path.SelectedPath);
            FileInfo[] Files = dInfo.GetFiles("*.xml");
            var fileNamesWithoutExtension = Files.Select(fi => Path.GetFileNameWithoutExtension(fi.Name));
            foreach(string file in fileNamesWithoutExtension)
            {
                Column_FB_Selection.Items.Add(file);
            }
           
            txt_FB_SElectedPath.Text = _FBD_Path.SelectedPath;
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.dataGridView1.Rows[e.RowIndex].Selected = true;
                this.rowIndex = e.RowIndex;
                this.dataGridView1.CurrentCell = this.dataGridView1.Rows[e.RowIndex].Cells[1];
                this.contextMenuStrip1.Show(this.dataGridView1, e.Location);
                contextMenuStrip1.Show(Cursor.Position);
            }
        }



        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
            if (!this.dataGridView1.Rows[this.rowIndex].IsNewRow)
            {
                this.dataGridView1.Rows.RemoveAt(this.rowIndex);
            }
        }

        private void button_Create_Click(object sender, EventArgs e)
        {
            string MainBlock = (_FBD_Path.SelectedPath + "\\Empty_FC.xml");
            int arrLenght = dataGridView1.Rows.Count - 1;
            Mechanism[] Mechanisms = new Mechanism[arrLenght];
            progressBar1.Value = 0;

            
            for(int i = 0; i < arrLenght; i++)
            {
                Mechanisms[i].name = dataGridView1.Rows[i].Cells[0].Value.ToString();
                Mechanisms[i].type = dataGridView1.Rows[i].Cells[1].Value.ToString();
            }


            XDocument doc = XDocument.Load(MainBlock);
            doc.Root.Element("SW.Blocks.FC").Element("ObjectList").DescendantsAndSelf("SW.Blocks.CompileUnit").Remove();
            //Console.WriteLine(doc);

            progressBar1.PerformStep();

            int id = Mechanisms.Length * 10 + 90;
            for(int i = 0; i < Mechanisms.Length; i++)
            {
                XDocument doc_to_Extract = XDocument.Load(_FBD_Path.SelectedPath + "\\Template_FC.xml");// + Mechanisms[i].type);"SW.Blocks.CompileUnit"                                                                                              //XElement ele_NewFragment = doc_to_Extract.Root.Element("SW.Blocks.FC").Element("ObjectList").Descendants().Where(z => z.Elements("SW.Blocks.CompileUnit").Any (x => (string)x.Attribute("Name") == a ));
                XElement ele_NewFragment = doc_to_Extract.Root.Element("SW.Blocks.FC").Element("ObjectList").Elements("SW.Blocks.CompileUnit").Where(d => d.Descendants().Any(s => (string)s.Attribute("Name") == Mechanisms[i].type)).FirstOrDefault();
                //Console.WriteLine(ele_NewFragment+"\n");

                //Extract ID's in order to update them
                IEnumerable<XElement> extract_ID = from z in ele_NewFragment.DescendantsAndSelf()
                                                   where (string)z.Attribute("ID") != null
                                                   select z;
                foreach(XElement Xid in extract_ID)
                {
                    Xid.Attribute("ID").Value = id.ToString();
                    id--;
                }

                // Update name of instance DB according to array of mechanism names

                XNamespace SIE = "http://www.siemens.com/automation/Openness/SW/NetworkSource/FlgNet/v1";
                XElement extract_Name = ele_NewFragment.Element("AttributeList").Element("NetworkSource").Element(SIE +"FlgNet").Element(SIE + "Parts").Element(SIE + "Call").Element(SIE + "CallInfo").Element(SIE +"Instance").Element(SIE +"Component");
                extract_Name.Attribute("Name").Value = Mechanisms[i].name;
                //Console.WriteLine(extract_Name);
                doc.Root.Element("SW.Blocks.FC").Element("ObjectList").Elements().Where(x => (string)x.Attribute("CompositionName") == "Title").FirstOrDefault().AddBeforeSelf(ele_NewFragment);
                
            }

            doc.Save(@"C:\Users\V14\Desktop\Sample\Generated.xml");


            progressBar1.PerformStep();


            for(int i=0; i < Mechanisms.Length; i++)
            {
                CreateInstDB(Mechanisms[i].name, Mechanisms[i].type, i + 5);
            }

            progressBar1.PerformStep();

            for (int i = 0; i < Mechanisms.Length; i++)
            {
                string file = _FBD_Path.SelectedPath + "\\" + Mechanisms[i].name + ".xml";

                FindPlc();
                ImportBlocks(_controller, file);
            }

            progressBar1.PerformStep();

            string path_FC = _FBD_Path.SelectedPath + "\\" + "Generated" + ".xml";
            FindPlc();
            ImportBlocks(_controller, path_FC);
            progressBar1.PerformStep();

            MessageBox.Show("Done!");

        }

        private void CreateInstDB(string MechanismName,string ObjType, int InstDatablockNumber)
        {
            try
            {
                XDocument xDoc = XDocument.Load(_FBD_Path.SelectedPath + "\\" + "Inst_" + ObjType + ".xml");
                Console.WriteLine(xDoc);
                
                xDoc.Root.Element("SW.Blocks.InstanceDB").Element("AttributeList").SetElementValue("Name", MechanismName);
                
                xDoc.Root.Element("SW.Blocks.InstanceDB").Element("AttributeList").SetElementValue("Number", InstDatablockNumber.ToString());

                xDoc.Save(_FBD_Path.SelectedPath +"\\"+ MechanismName + ".xml");
            }
            catch(FileNotFoundException e)
            {
                MessageBox.Show(e.Message);
            }


        }

    }





}

