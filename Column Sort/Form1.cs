using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RobotOM;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;

namespace Column_Sort
{
    // made by Sam
    public partial class Form1 : Form
    {
        public RobotApplication robapp;
        private bool[,] panels;
        private int sizeX = 9 , sizeY = 9;
        public Form1()
        {
            InitializeComponent();
             
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                robapp = new RobotApplication();


                Generate();
            }
            catch (Exception E)
            {
                textBox1.AppendText(E.ToString());
            }
            finally
            {
                robapp = null;
            }

        }

        public void Generate()
        {
            panels = new bool[sizeX, sizeY];

            Random rnd = new Random();
            int i = rnd.Next(0, sizeX);
            int j = rnd.Next(0, sizeY);
            int[] directionsX = { 0, 1, 0, -1 };
            int[] directionsY = { 1, 0, -1, 0 };

            while (ContainsCoordinates(j, i) && panels[j, i] == false)
            {
                CreatePanelAtPoint(j, i);
                panels[j, i] = true;
                textBox1.AppendText($"{j},{i} created {Environment.NewLine}");
                
                int rand = rnd.Next(0, 3);
                i += directionsX[rand];
                j += directionsY[rand];
            }
            textBox1.AppendText($"Ended at {j},{i} {Environment.NewLine}");
        }

        public bool ContainsCoordinates(float x, float y)
        {
            return x >= 0 && x < sizeX && y >= 0 && y < sizeY;
        }

        public void CreatePanelAtPoint(float x, float y)
        {
            //Create Nodes
            IRobotCollection robotNodeServer = robapp.Project.Structure.Nodes.GetAll();
            int totalnodes = robotNodeServer.Count + 1;
            int count = 0;
            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    robapp.Project.Structure.Nodes.Create(totalnodes + count, x + i, y + j, 0);
                    count++;
                }
            }

            //Create Array from nodes
            RobotNumbersArray robotNumbersArray = robapp.CmpntFactory.Create(IRobotComponentType.I_CT_NUMBERS_ARRAY);
            robotNumbersArray.SetSize(4);
            int[] offsetCount = { 0, 1, 3, 2 };
            for (int i = 0; i < 4; i++)
            {
                robotNumbersArray.Set(i+1, totalnodes + offsetCount[i]);
            }

            //Create finite element
            int totalFE = robapp.Project.Structure.FiniteElems.GetAll().Count + 1;
            robapp.Project.Structure.FiniteElems.Create(totalFE, robotNumbersArray);

            //Create panel
            int totalObjects = robapp.Project.Structure.Objects.GetAll().Count + 1;
            robapp.Project.Structure.Objects.CreateOnFiniteElems(totalFE.ToString(), totalObjects);
            IRobotObjObject panel = robapp.Project.Structure.Objects.Get(totalObjects) as IRobotObjObject;
            panel.SetLabel(IRobotLabelType.I_LT_PANEL_THICKNESS, "TH30_CONCR");
            panel.SetLabel(IRobotLabelType.I_LT_PANEL_CALC_MODEL, "Shell");
        }

        public void GetAndDisplayNodeLocationsOnGraph()
        {
            RobotApplication robotApplication = new RobotApplication();
            chart1.Series.Clear();

            var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "Series1",
                Color = System.Drawing.Color.Green,
                IsVisibleInLegend = false,
                ChartType = SeriesChartType.Point
            };

            this.chart1.Series.Add(series1);
            try
            {
                textBox1.AppendText("Method Running" + Environment.NewLine);
                RobotStructureCache robotStructureCache = robotApplication.Project.Structure.CreateCache();
                IRobotCollection robotNodeCollection = robotApplication.Project.Structure.Nodes.GetAll();

                int[] nodes = new int[100];
                int count = 0;

                for (float i = 0; i < 10; i++)
                {
                    for (float j = 0; j < 10; j++)
                    {
                        nodes[count] = DoesNodeExistAtXY(j, i, robotNodeCollection.Count, robotStructureCache);
                        bool nodeReal = nodes[count] != 0;
                        if (nodeReal)
                        {
                            series1.Points.AddXY(j, i);
                        }
                        textBox1.AppendText($"{j},{i} node is {nodeReal.ToString()} {Environment.NewLine}");
                        count++;

                    }
                }
                chart1.Invalidate();


            }
            catch (Exception e)
            {
                textBox1.Text = e.ToString();
            }
            finally
            {
                robotApplication = null;
            }

        }

        public void UpdateMeshesForCorrectCalculationSettings()
        {
            RobotApplication robotApplication = new RobotApplication();
            try
            {
                RobotProjectPreferences ProjectPrefs = robotApplication.Project.Preferences;
                RobotMeshParams MeshParams = robotApplication.Project.Preferences.MeshParams;
                MeshParams.SurfaceParams.Method.Method = IRobotMeshMethodType.I_MMT_COONS;
                MeshParams.SurfaceParams.Generation.Type = IRobotMeshGenerationType.I_MGT_ELEMENT_SIZE;
                MeshParams.SurfaceParams.Generation.ElementSize = 1;
                MeshParams.SurfaceParams.Coons.PanelDivisionType = IRobotMeshPanelDivType.I_MPDT_SQUARE_IN_RECT;
                MeshParams.SurfaceParams.Coons.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;
                MeshParams.SurfaceParams.FiniteElems.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;

                RobotSelection robotSelection = robotApplication.Project.Structure.Selections.Create(IRobotObjectType.I_OT_PANEL);
                robotSelection.FromText("all");

                RobotObjObjectCollection panelCol = robotApplication.Project.Structure.Objects.GetMany(robotSelection) as RobotObjObjectCollection;

                for (int i = 1; i <= panelCol.Count; i++)
                {
                    RobotObjObject panel = panelCol.Get(i);
                    panel.Main.Attribs.Meshed = 1;
                    panel.Update();
                    panel.Mesh.Generate();
                }
            }
            catch (Exception e)
            {
                textBox1.Text = e.ToString();
            }
            finally
            {
                robotApplication = null;
            }

        }

        

        public int DoesNodeExistAtXY(float x, float y, int totalNodes, RobotStructureCache robotStructureCache)
        {
            int i = robotStructureCache.EnsureNodeExist(x, y, 0);
            try
            {
                IRobotDataObject node = robapp.Project.Structure.Nodes.Get(i);
                if(node.Number == i)
                {
                    return 1;
                }
            }
            catch(Exception e)
            {

            }
            finally
            {
                
            }
            return 0;
        }

        #region old code
        public void GetCases()
        {

            RobotApplication robapp = new RobotApplication();

            if (robapp == null)
            {
                return;
            }

            robapp = null;

            IRobotCaseCollection robotCaseCollection = robapp.Project.Structure.Cases.GetAll();
            for (int i = 0; i < robotCaseCollection.Count; i++)
            {
                try
                {
                    IRobotCase robotCase = robapp.Project.Structure.Cases.Get(i);
                    if (robotCase != null)
                    {
                        listBox1.Items.Add(robotCase.Name);
                    }
                }
                catch (Exception e)
                {

                }
            }

            if (listBox1.SelectedIndex == -1)
            {
                listBox1.SelectedIndex = 0;
            }



        }

        public void CopyData()
        {
            button1.Text = "Loading...";
            //this.WindowState = FormWindowState.Minimized;


            //-------------------------------------
            //Load Cases
            IRobotCaseCollection robotCaseCollection = robapp.Project.Structure.Cases.GetAll();
            int loadCase = 0;
            int FindCase(string casename)
            {
                int number = 1;
                IRobotCase robotCase = robapp.Project.Structure.Cases.Get(1);
                for (int i = 0; i < robotCaseCollection.Count; i++)
                {
                    robotCase = robapp.Project.Structure.Cases.Get(i);
                    if (robotCase != null)
                    {
                        if (robotCase.Name == casename)
                        {
                            number = i;
                            break;
                        }
                    }
                }
                loadCase = number;
                return number;

            }


            //-------------------------------------
            //Get Number of Bars Selected
            RobotSelection barSel = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_BAR);
            //Get All Load Cases
            RobotSelection casSel = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_CASE);

            //Get Bar and Node Data
            IRobotBarServer robotBarServer = robapp.Project.Structure.Bars;
            IRobotNodeServer inds = robapp.Project.Structure.Nodes;

            //Get a List of the bars and Setup bar information Struct
            int[] barSelArray = new int[barSel.Count];
            BeamDataStruct[] beamData = new BeamDataStruct[barSelArray.Length];
            for (int i = 1; i < barSel.Count + 1; i++)
            {
                //Setup bar no. array
                barSelArray[i - 1] = barSel.Get(i);

                //Get node information from bar data
                IRobotBar bar = (IRobotBar)robotBarServer.Get(barSelArray[i - 1]);
                int startNodeNo = bar.StartNode;
                int endNodeNo = bar.EndNode;
                IRobotNode startNode = (IRobotNode)inds.Get(startNodeNo);
                IRobotNode endNode = (IRobotNode)inds.Get(endNodeNo);

                //If a Beam, Skip
                if (startNode.Z == endNode.Z) { continue; }

                //Which is highest node
                IRobotNode node = (startNode.Z > endNode.Z) ? startNode : endNode;

                //Populate beam data from node and bar data.
                beamData[i - 1].barNo = barSelArray[i - 1];

                IRobotBarSectionData sectData = bar.GetLabel(IRobotLabelType.I_LT_BAR_SECTION).Data;
                double depth = sectData.GetValue(IRobotBarSectionDataValue.I_BSDV_BF);
                double breath = sectData.GetValue(IRobotBarSectionDataValue.I_BSDV_D);
                if (depth < breath)
                {
                    double holder = breath;
                    breath = depth;
                    depth = holder;
                }
                depth = depth * 1000;
                breath = breath * 1000;
                beamData[i - 1].section = $"C1 {depth} x {breath}";
                beamData[i - 1].x = node.X;
                beamData[i - 1].y = node.Y;
                beamData[i - 1].z = node.Z;
                beamData[i - 1].height = bar.Length;
                IRobotMaterialData concrete = bar.GetLabel(IRobotLabelType.I_LT_MATERIAL).Data;
                Double concreteS = concrete.RE / 1000000;
                beamData[i - 1].concreteStrength = concreteS.ToString();

            }

            textBox1.AppendText("\r\nSorting\r\n");
            beamData = beamData.OrderBy(x => x.z).ToArray();
            beamData = beamData.OrderBy(x => x.y).ToArray();
            beamData = beamData.OrderBy(x => x.x).ToArray();

            int group = 1;
            int posInGroup = 0;
            for (int i = 0; i < beamData.Length; i++)
            {
                posInGroup = 0;


                for (int j = 0; j < beamData.Length; j++)
                {
                    if (beamData[i].x - beamData[j].x < 0.0001 && beamData[i].y - beamData[j].y < 0.0001 && beamData[i].barNo != beamData[j].barNo)
                    {
                        if (beamData[j].group != 0)
                        {
                            beamData[i].group = beamData[j].group;
                            for (int k = 0; k < beamData.Length; k++)
                            {
                                if (beamData[i].group == beamData[k].group && beamData[i].barNo != beamData[k].barNo)
                                {
                                    posInGroup++;
                                }
                            }
                            beamData[i].posInGroup = posInGroup;
                        }
                        else
                        {
                            beamData[i].group = group;
                            group++;
                        }
                        break;
                    }
                }
            }

            void CalculateResults()
            {
                textBox1.AppendText($"\r\nStarting calculation: {DateTime.Now.ToString("h:mm:ss tt")}");
                RobotExtremeParams robotExtremeParams = robapp.CmpntFactory.Create(IRobotComponentType.I_CT_EXTREME_PARAMS);
                robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_CASE).FromText(FindCase(listBox1.SelectedItem.ToString()).ToString());
                robotExtremeParams.Selection.Set(IRobotObjectType.I_OT_CASE, casSel);
                IRobotBarForceServer robotBarResultServer = robapp.Project.Structure.Results.Bars.Forces;
                int total = beamData.Length;
                bool firstLoop = true;
                string columnsSelected = "";
                for (int i = 0; i < beamData.Length; i++)
                {
                    DateTime startTime = DateTime.Now;
                    textBox1.AppendText($"\r\nStart Calculation {i + 1} / {total} before bar selection: {DateTime.Now.ToString("h:mm:ss tt")}");
                    robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_BAR).FromText(beamData[i].barNo.ToString());
                    robotExtremeParams.Selection.Set(IRobotObjectType.I_OT_BAR, barSel);


                    //MZ
                    robotExtremeParams.ValueType = IRobotExtremeValueType.I_EVT_FORCE_BAR_MZ;

                    if (Math.Abs(robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).Value) > Math.Abs(robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).Value))
                    {
                        beamData[i].mZForceServer = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).CaseCmpnt, 1);
                        beamData[i].mZForceServerbtm = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).CaseCmpnt, 0);

                    }
                    else
                    {
                        beamData[i].mZForceServer = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).CaseCmpnt, 1);
                        beamData[i].mZForceServerbtm = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).CaseCmpnt, 0);
                    }

                    beamData[i].mzValue = Math.Abs(beamData[i].mZForceServer.MZ) > Math.Abs(beamData[i].mZForceServerbtm.MZ) ? beamData[i].mZForceServer.FX : beamData[i].mZForceServerbtm.FX;



                    //MY
                    robotExtremeParams.ValueType = IRobotExtremeValueType.I_EVT_FORCE_BAR_MY;

                    if (Math.Abs(robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).Value) > Math.Abs(robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).Value))
                    {
                        beamData[i].mYForceServer = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).CaseCmpnt, 1);
                        beamData[i].mYForceServerbtm = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).CaseCmpnt, 0);
                    }
                    else
                    {
                        beamData[i].mYForceServer = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).CaseCmpnt, 1);
                        beamData[i].mYForceServerbtm = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).CaseCmpnt, 0);
                    }

                    beamData[i].myValue = Math.Abs(beamData[i].mYForceServer.MY) > Math.Abs(beamData[i].mYForceServerbtm.MY) ? beamData[i].mYForceServer.FX : beamData[i].mYForceServerbtm.FX;



                    //FX
                    robotExtremeParams.ValueType = IRobotExtremeValueType.I_EVT_FORCE_BAR_FX;

                    if (Math.Abs(robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).Value) > Math.Abs(robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).Value))
                    {
                        beamData[i].fXForceServer = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).CaseCmpnt, 1);
                        beamData[i].fXForceServerbtm = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MaxValue(robotExtremeParams).CaseCmpnt, 0);
                    }
                    else
                    {
                        beamData[i].fXForceServer = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).CaseCmpnt, 1);
                        beamData[i].fXForceServerbtm = robotBarResultServer.ValueEx(beamData[i].barNo, loadCase, robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).CaseCmpnt, 0);
                    }

                    beamData[i].fxValue = Math.Abs(beamData[i].fXForceServer.FX) > Math.Abs(beamData[i].fXForceServerbtm.FX) ? beamData[i].fXForceServer.FX : beamData[i].fXForceServerbtm.FX;


                    double totalTime = (DateTime.Now - startTime).TotalSeconds;
                    textBox1.AppendText($"\r\nEnd Calculation {i + 1} / {total} {DateTime.Now.ToString("h:mm:ss tt")} \r\nTime taken: {totalTime}");

                    if (firstLoop)
                    {
                        textBox1.AppendText($"\r\nEstimated finish time: {DateTime.Now.AddSeconds(total * totalTime).ToString("h:mm:ss tt")}");
                        firstLoop = false;
                    }

                    columnsSelected += $"{beamData[i].barNo.ToString()} ";
                }

                textBox1.AppendText($"\r\ncolumns selected {columnsSelected}");
                robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_BAR).FromText(columnsSelected);
            }

            int maxCol = 1;

            void WriteResults()
            {


                int column = 1;
                int currentGroup = 0;
                for (int i = 0; i < beamData.Length; i++)
                {
                    if (beamData[i].group == currentGroup)
                    {
                        column = beamData[i].posInGroup;
                        if (column >= maxCol) { maxCol = column; }
                    }
                    else
                    {
                        currentGroup++;
                        column = 0;
                    }

                    int row = currentGroup + 2;
                    int columnPos = beamData[i].posInGroup + 1;
                    WriteCell(row, 0, columnPos.ToString());

                    WriteCell(row, 1 + 22 * column, beamData[i].section.ToString());
                    WriteCell(row, 2 + 22 * column, beamData[i].barNo.ToString());
                    WriteCell(row, 3 + 22 * column, beamData[i].concreteStrength.ToString());
                    WriteCell(row, 4 + 22 * column, beamData[i].group.ToString());
                    WriteCell(row, 5 + 22 * column, beamData[i].posInGroup.ToString());
                    WriteCell(row, 6 + 22 * column, beamData[i].height.ToString());

                    WriteCell(row, 7 + 22 * column, (beamData[i].fxValue / 1000).ToString());
                    WriteCell(row, 8 + 22 * column, (beamData[i].fXForceServer.MY / 1000).ToString());
                    WriteCell(row, 9 + 22 * column, (beamData[i].fXForceServer.MZ / 1000).ToString());
                    WriteCell(row, 10 + 22 * column, (beamData[i].fXForceServerbtm.MY / 1000).ToString());
                    WriteCell(row, 11 + 22 * column, (beamData[i].fXForceServerbtm.MZ / 1000).ToString());


                    WriteCell(row, 12 + 22 * column, (beamData[i].mzValue / 1000).ToString());
                    WriteCell(row, 13 + 22 * column, (beamData[i].mZForceServer.MY / 1000).ToString());
                    WriteCell(row, 14 + 22 * column, (beamData[i].mZForceServer.MZ / 1000).ToString());
                    WriteCell(row, 15 + 22 * column, (beamData[i].mZForceServerbtm.MY / 1000).ToString());
                    WriteCell(row, 16 + 22 * column, (beamData[i].mZForceServerbtm.MZ / 1000).ToString());


                    WriteCell(row, 17 + 22 * column, (beamData[i].myValue / 1000).ToString());
                    WriteCell(row, 18 + 22 * column, (beamData[i].mYForceServer.MY / 1000).ToString());
                    WriteCell(row, 19 + 22 * column, (beamData[i].mYForceServer.MZ / 1000).ToString());
                    WriteCell(row, 20 + 22 * column, (beamData[i].mYForceServerbtm.MY / 1000).ToString());
                    WriteCell(row, 21 + 22 * column, (beamData[i].mYForceServerbtm.MZ / 1000).ToString());

                }

                WriteCell(0, 0, currentGroup.ToString());
            }

            void PopulateHeaders()
            {
                for (int i = 0; i <= maxCol; i++)
                {
                    //Headers
                    WriteCell(1, 1 + 22 * i, "Cross Section");
                    WriteCell(1, 2 + 22 * i, "Bar No.");
                    WriteCell(1, 3 + 22 * i, "Concrete Strength");
                    WriteCell(1, 4 + 22 * i, "Group");
                    WriteCell(1, 5 + 22 * i, "Pos In Group");
                    WriteCell(1, 6 + 22 * i, "Length");

                    //FZ Max
                    WriteCell(1, 7 + 22 * i, "Fx (Max) [kN]");
                    WriteCell(1, 8 + 22 * i, "My (Top) [kNm]");
                    WriteCell(1, 9 + 22 * i, "Mz (Top) [kNm]");
                    WriteCell(1, 10 + 22 * i, "My (Btm) [kNm]");
                    WriteCell(1, 11 + 22 * i, "Mz (Btm) [kNm]");

                    //MX Max
                    WriteCell(1, 12 + 22 * i, "Fx (Max) [kN]");
                    WriteCell(1, 13 + 22 * i, "My (Top) [kNm]");
                    WriteCell(1, 14 + 22 * i, "Mz (Top) [kNm]");
                    WriteCell(1, 15 + 22 * i, "My (Btm) [kNm]");
                    WriteCell(1, 16 + 22 * i, "Mz (Btm) [kNm]");

                    //MY Max
                    WriteCell(1, 17 + 22 * i, "Fx (Max) [kN])");
                    WriteCell(1, 18 + 22 * i, "My (Top) [kNm]");
                    WriteCell(1, 19 + 22 * i, "Mz (Top) [kNm]");
                    WriteCell(1, 20 + 22 * i, "My (Btm) [kNm]");
                    WriteCell(1, 21 + 22 * i, "Mz (Btm) [kNm]");

                    //Headers
                    WriteCell(0, 9 + 22 * i, "Fx (Max)");
                    WriteCell(0, 14 + 22 * i, "Mz (Max)");
                    WriteCell(0, 19 + 22 * i, "My (Max)");
                }
            }


            WriteData();
            CalculateResults();
            WriteResults();
            PopulateHeaders();
            SaveExcel();
            CloseExcel();
            button1.Text = "Start";
            textBox1.AppendText("\r\nDone, view your documents for the file named 'Results for bars ~date~', you may close this window or select more columns and press 'Start'.");
            robapp = null;
            this.WindowState = FormWindowState.Normal;

        }



        struct BeamDataStruct
        {
            //Bar Data
            public int barNo;
            public string section;
            public double x;
            public double y;
            public double z;
            public double height;
            public string concreteStrength;


            //Sorting Data
            public int group;
            public int posInGroup;

            //Force Data
            public IRobotBarForceData mZForceServer;
            public IRobotBarForceData mZForceServerbtm;
            public IRobotBarForceData mYForceServer;
            public IRobotBarForceData mYForceServerbtm;
            public IRobotBarForceData fXForceServer;
            public IRobotBarForceData fXForceServerbtm;
            public double fxValue;
            public double myValue;
            public double mzValue;

        };



        string path = "";
        _Application excel = new Excel.Application();
        Workbook wb;
        Worksheet ws;

        public void OpenExcel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;

            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return "";
        }

        public void WriteCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }

        public void SaveExcel()
        {
            wb.Save();
        }

        public void CloseExcel()
        {
            if (wb != null)
            {
                wb.Close();
                wb = null;
            }

        }

        public void WriteData()
        {
            string FileName = "Results for Bars " + DateTime.Now.ToString() + " .xlsx";
            FileName = FileName.Replace(":", "-");
            FileName = FileName.Replace("/", "-");
            wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            wb.SaveAs(@FileName);
            OpenExcel(@FileName, 1);

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            textBox1.Text = $"\r\nSelected Load Case: {listBox1.SelectedItem.ToString()} \r\nSelect some columns in robot and press 'Start'.";

        }
    }
    #endregion

}
