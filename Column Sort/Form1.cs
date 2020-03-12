using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RobotOM;
//using Microsoft.Office.Interop.Excel;
//using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;
using static Tensorflow.Binding;
using System.Numerics;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Threading;

namespace Column_Sort
{
    // made by Sam
    public partial class Form1 : Form
    {
        public RobotApplication robapp;
        private bool[,] panels;
        private int sizeX = 8, sizeY = 8;
        //private int numOfColumns;
        public double DeflectionThreshold = -5.65; //sqrt 2 *1000/ 250
        public double ColumnFactor = 10;

        int lastNode = 0;
        int lastObject = 0;
        int numOfTiles;
        DateTime StartTime;
        HashSet<int> AllNodesInModel;
        Dictionary<Vector2, int> NodeLocationsAndNumbersDict;
        int runs = 0;
        string columnSection = "250x250_Column";
        string panelSection = "TH20_C20";

        private Child PassdownChild;
        public Form1()
        {
            InitializeComponent();
            PassdownChild = null;
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
                robapp.Interactive = 0;
                robapp.UserControl = false;
                int samples = int.Parse(textBox3.Text.ToString());
                int gens = int.Parse(textBox4.Text.ToString());

                if (!checkBox1.Checked)
                {
                    LoadData();
                }

                if (PassdownChild != null)
                {
                    textBox1.Text = $"Gen: {PassdownChild.Gen}, start.";
                }

                StartTime = DateTime.Now;
                //double timer = 60 * 10;
                do
                {
                    InternalTest(samples);
                    textBox1.AppendText($"Gen: {PassdownChild.Gen}, next.");
                    //timer -= (DateTime.Now - StartTime).TotalSeconds;
                    //textBox1.AppendText(timer.ToString() + Environment.NewLine);
                    //StartTime = DateTime.Now;
                    runs++;
                } while ((PassdownChild.Gen <= gens));

            }
            catch (Exception E)
            {
                PrintErrorMessage(E);
            }
            finally
            {
                if (PassdownChild != null)
                {
                    SaveData(PassdownChild);
                }
                robapp.Visible = 1;
                robapp.Interactive = 1;
                robapp.UserControl = true;
                robapp = null;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                textBox1.AppendText(ReadAllTextFiles());
            }
            catch (Exception E)
            {
                PrintErrorMessage(E);
            }
            finally
            {

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                robapp = new RobotApplication();
                Screenshot(10, new Vector2(10, 10));
            }
            catch (Exception E)
            {
                PrintErrorMessage(E);
            }
            finally
            {
                robapp = null;
            }
        }

        

        void InternalTest(int samples)
        {
            DateTime startTime = DateTime.Now;

            DeleteAll(runs);

            //Create a floor plan

            Random rnd = new Random();
            numOfTiles = rnd.Next(4, 14);

            List<Vector2> nodeLocations = GenerateNumofFloorTiles(numOfTiles);

            //Copy the plan
            int tiles = samples;

            CopyGeneratedTiles(tiles - 1);

            //Analyse the samples
            double[] inputs = GetAndDisplayNodeLocationsOnGraph();

            SetTopNode();
            NodeLocationsAndNumbersDict = NodeLocationsAndNumbers(nodeLocations, tiles);
            AllNodesInModel = GetAllNodes();

            //Create NN
            List<Child> Children = new List<Child>();
            int k = 0;
            for (int i = 0; i < tiles; i++)
            {
                for (int j = 0; j < tiles; j++)
                {
                    if (PassdownChild == null)
                    {
                        Children.Add(ConstructNN(k, inputs, new Vector2(15 * i, 15 * j)));
                    }
                    else
                    {
                        Children.Add(ConstructNN(k, inputs, new Vector2(15 * i, 15 * j), PassdownChild));
                    }

                    k += 1;
                }
            }

            Calculate();



            k = 0;
            for (int i = 0; i < tiles; i++)
            {
                for (int j = 0; j < tiles; j++)
                {
                    Children[k].Score = GetResults(nodeLocations, new Vector2(15 * i, 15 * j), Children[k].NumberOfColumns);
                    k += 1;
                }
            }


            List<Child> orderedChildren = Children.OrderBy(o => o.Score).ToList();
            List<double> scoresList = new List<double>();
            foreach (var item in orderedChildren)
            {
                scoresList.Add(item.Score);
            }

            PassdownChild = orderedChildren[0];

            SaveScoresToTextFile(PassdownChild.Gen, scoresList, numOfTiles, PassdownChild.NumberOfColumns);

            PassdownChild.Gen += 1;
            textBox1.AppendText($"Gen: {PassdownChild.Gen - 1}, runtime: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

        }

        private void CreateMaterialsColumnsAndPanels()
        {
            //Create a 250x250 Column
            //Apply C20 material
            //Create a 40mm Panel
            //Apply C20 material

            var columnSectionLabel = robapp.Project.Structure.Labels.Create(IRobotLabelType.I_LT_BAR_SECTION, columnSection);
            RobotBarSectionData robotBarSectionData = columnSectionLabel.Data;
            robotBarSectionData.ShapeType = IRobotBarSectionShapeType.I_BSST_CONCR_COL_R;
            robotBarSectionData.MaterialName = "C15";
            robotBarSectionData.Concrete.SetValue(IRobotBarSectionConcreteDataValue.I_BSCDV_COL_B, 0.25);
            robotBarSectionData.Concrete.SetValue(IRobotBarSectionConcreteDataValue.I_BSCDV_COL_H, 0.25);

            robotBarSectionData.SetValue(IRobotBarSectionDataValue.I_BSDV_BF, 1);
            robotBarSectionData.SetValue(IRobotBarSectionDataValue.I_BSDV_D, 1);
            robotBarSectionData.CalcNonstdGeometry();
            robapp.Project.Structure.Labels.Store(columnSectionLabel);


            var panelSectionLabel = robapp.Project.Structure.Labels.Create(IRobotLabelType.I_LT_PANEL_THICKNESS, panelSection);
            RobotThicknessData robotPanelSectionData = panelSectionLabel.Data;
            robotPanelSectionData.ThicknessType = IRobotThicknessType.I_TT_HOMOGENEOUS;
            robotPanelSectionData.MaterialName = "C15";
            IRobotThicknessHomoData robotThicknessHomoData = robotPanelSectionData.Data;
            robotThicknessHomoData.ThickConst = 0.04;
            robapp.Project.Structure.Labels.Store(panelSectionLabel);
        }

        void SaveScoresToTextFile(int gen, List<double> scores, int tiles, int columns)
        {
            string FileName = $"{DateTime.Now.ToString()} , Gen {gen.ToString("0000")}.txt";
            FileName = FileName.Replace(":", "-");
            FileName = FileName.Replace("/", "-");

            var currentDirectory = Path.GetDirectoryName(Assembly.GetCallingAssembly().Location);
            currentDirectory = currentDirectory.Substring(0, currentDirectory.Length - 22) + @"\Export\";

            string data = gen.ToString() + "," + scores.Average().ToString("G17") + "," + StandardDeviation(scores).ToString("G17") + "," + tiles.ToString() + "," + columns.ToString();
            foreach (var score in scores)
            {
                data += "," + score.ToString("G17");
            }
            File.WriteAllText(currentDirectory + FileName, data);
        }

        string ReadAllTextFiles()
        {
            string text = "";

            var currentDirectory = Path.GetDirectoryName(Assembly.GetCallingAssembly().Location);
            currentDirectory = currentDirectory.Substring(0, currentDirectory.Length - 22) + @"\Export\Test\";

            DirectoryInfo d = new DirectoryInfo(currentDirectory);//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.txt"); //Getting Text files
            foreach (FileInfo file in Files)
            {
                text += File.ReadAllText(file.FullName) + Environment.NewLine;
            }

            return text;
        }

        double StandardDeviation(IEnumerable<double> values)
        {
            double avg = values.Average();
            return Math.Sqrt(values.Average(v => Math.Pow(v - avg, 2)));
        }


        void Screenshot(int gen, Vector2 bestLocation)
        {
            //Setup wide angle view

            robapp.UserControl = true;
            robapp.Interactive = 1;
            robapp.Visible = 1;
            robapp.Window.Activate();
            IRobotView3 robotView = robapp.Project.ViewMngr.GetView(1) as IRobotView3;
            robotView.Redraw(1);
            robotView.Projection = IRobotViewProjection.I_VP_3DXYZ;
            robotView.Rotate(IRobotGeoCoordinateAxis.I_GCA_OZ, -1.0472);
            robotView.Rotate(IRobotGeoCoordinateAxis.I_GCA_OY, 1.22);
            robotView.Rotate(IRobotGeoCoordinateAxis.I_GCA_OZ, 1.5708);
            robapp.Project.ViewMngr.Refresh();

            RobotViewScreenCaptureParams robotViewScreenCaptureParams = robapp.CmpntFactory.Create(IRobotComponentType.I_CT_VIEW_SCREEN_CAPTURE_PARAMS);
            string screenShotName = $"{DateTime.Now.ToString()}, Gen {gen}, Columns";
            robotViewScreenCaptureParams.Name = screenShotName;
            robotViewScreenCaptureParams.UpdateType = IRobotViewScreenCaptureUpdateType.I_SCUT_CURRENT_VIEW;
            robotViewScreenCaptureParams.Resolution = IRobotViewScreenCaptureResolution.I_VSCR_4096;
            robotView.MakeScreenCapture(robotViewScreenCaptureParams);
            robapp.Project.PrintEngine.AddScToReport(screenShotName);

            robotView.ParamsFeMap.CurrentResult = IRobotViewFeMapResultType.I_VFMRT_GLOBAL_DISPLACEMENT_Z;
            robapp.Project.ViewMngr.Refresh();
            screenShotName = $"{DateTime.Now.ToString()}, Gen {gen}, Deflections";
            robotViewScreenCaptureParams.Name = screenShotName;
            robotView.MakeScreenCapture(robotViewScreenCaptureParams);
            robapp.Project.PrintEngine.AddScToReport(screenShotName);

            robotView.SetZoom(bestLocation.X - 1, bestLocation.Y + 11, bestLocation.X + 11, bestLocation.Y - 1);
            robapp.Project.ViewMngr.Refresh();
            screenShotName = $"{DateTime.Now.ToString()}, Gen {gen}, Best";
            robotViewScreenCaptureParams.Name = screenShotName;
            robotView.MakeScreenCapture(robotViewScreenCaptureParams);
            robapp.Project.PrintEngine.AddScToReport(screenShotName);

            string FileName = $"{DateTime.Now.ToString()} , Gen {gen.ToString("0000")}.rtf";
            FileName = FileName.Replace(":", "-");
            FileName = FileName.Replace("/", "-");

            var currentDirectory = Path.GetDirectoryName(Assembly.GetCallingAssembly().Location);
            currentDirectory = currentDirectory.Substring(0, currentDirectory.Length - 22) + @"\Export\";

            robapp.Project.PrintEngine.SaveReportToFile(currentDirectory + FileName, IRobotOutputFileFormat.I_OFF_RTF_JPEG);

            robapp.UserControl = false;
            robapp.Interactive = 0;
        }

        //TODO why does this call the does node exist method so much??
        double GetResults(List<Vector2> nodeLocations, Vector2 offset, int numOfColumns = 0)
        {
            HashSet<int> existingNodes = new HashSet<int>();
            foreach (var node in nodeLocations)
            {
                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        Vector2 nodeCoords = new Vector2(node.X + offset.X + i, node.Y + offset.Y + j);
                        int num = DoesNodeExistAtXY(nodeCoords, true);
                        existingNodes.Add(num);
                    }
                }
            }


            double deflectionScore = 0;
            double highestDeflectionZ = 0;

            foreach (var nodeNum in existingNodes)
            {
                double val = robapp.Project.Structure.Results.Nodes.Displacements.Value(nodeNum, 1).UZ * 1000;

                if (val > DeflectionThreshold)
                {
                    deflectionScore += val * 2;
                }

                highestDeflectionZ = Math.Min(highestDeflectionZ, val);
            }
            return -((deflectionScore + highestDeflectionZ - (numOfColumns * ColumnFactor)));
        }

        /*
          var allNodes = robapp.Project.Structure.Nodes.GetAll();
            double total = 0;
            textBox1.AppendText(allNodes.Count + Environment.NewLine);
            for (int i = 1; i <= allNodes.Count; i++)
            {
                total += robapp.Project.Structure.Results.Nodes.Displacements.Value(i, 1).UZ;
                
            }
            total *= 1000;
            textBox1.AppendText(total.ToString() + Environment.NewLine + numOfColumns.ToString());

         */

        /*
          var allNodes = robapp.Project.Structure.Nodes.GetAll();
            double total = 0;
            textBox1.AppendText(allNodes.Count + Environment.NewLine);


            RobotSelection nodeSel = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_NODE);
            RobotSelection casSel = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_CASE);
            nodeSel.FromText("all");
            casSel.FromText("all");

            RobotExtremeParams robotExtremeParams = robapp.CmpntFactory.Create(IRobotComponentType.I_CT_EXTREME_PARAMS);
            robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_CASE).FromText("DL1");
            robotExtremeParams.Selection.Set(IRobotObjectType.I_OT_CASE, casSel);
            IRobotBarForceServer robotBarResultServer = robapp.Project.Structure.Results.Bars.Forces;

            for (int i = 1; i <= allNodes.Count; i++)
            {
                total += robapp.Project.Structure.Results.Nodes.Displacements.Value(i, 1).UZ;

                IRobotDataObject node = allNodes.Get(i);

                robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_NODE).FromText(node.Number.ToString());
                robotExtremeParams.Selection.Set(IRobotObjectType.I_OT_NODE, nodeSel);
                robotExtremeParams.ValueType = IRobotExtremeValueType.I_EVT_DEFLECTION_UZ;
                total = robapp.Project.Structure.Results.Extremes.MinValue(robotExtremeParams).Value;
                textBox1.AppendText(total.ToString("F5") + Environment.NewLine); 
            }
            */

        public void SaveData(Child bestChild)
        {
            File.Create("data.txt").Close();

            using (StreamWriter writetext = new StreamWriter("data.txt"))
            {

                for (int i = 0; i < 8; i++)
                {
                    for (int j = 0; j < 100; j++)
                    {
                        for (int k = 0; k < 100; k++)
                        {
                            writetext.Write(bestChild.weights[i, j, k].ToString("G17") + "\n");
                        }
                    }
                }

                writetext.Write(":\n");

                for (int i = 0; i < 8; i++)
                {
                    for (int j = 0; j < 100; j++)
                    {
                        writetext.Write(bestChild.bias[i, j].ToString("G17") + "\n");
                    }
                }

                writetext.Write(":\n");

                writetext.Write(bestChild.Gen);
            }

        }

        public void LoadData()
        {
            string text = "";
            using (StreamReader readtext = new StreamReader("data.txt"))
            {
                text = readtext.ReadToEnd();
            }

            if (text != "")
            {
                PassdownChild = new Child(0, new Vector2(0, 0), 0);
            }
            else
            {
                return;
            }

            string[] splitText = text.Split(':');

            string weights = splitText[0];
            string bias = splitText[1];
            string generation = splitText[2];

            int p = 0;
            int q = 0;

            string[] allWeights = weights.Split('\n');
            List<string> allBias = bias.Split('\n').ToList<string>();

            allBias.RemoveAt(0);
            allBias.RemoveAt(allBias.Count - 1);

            PassdownChild.Gen = int.Parse(generation);

            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        PassdownChild.weights[i, j, k] = double.Parse(allWeights[p]);
                        p++;
                    }
                    PassdownChild.bias[i, j] = float.Parse(allBias[q]);
                    q++;
                }
            }
        }

        public void Generate()
        {
            panels = new bool[sizeX, sizeY];

            Random rnd = new Random();
            int i = rnd.Next(0, sizeX);
            int j = rnd.Next(0, sizeY);
            int[] directionsX = { 0, 2, 0, -2 };
            int[] directionsY = { 2, 0, -2, 0 };

            while (ContainsCoordinates(new Vector2(j, i)))
            {
                if (panels[j, i] == false)
                {
                    CreatePanelAtPoint(new Vector2(j, i));
                    panels[j, i] = true;
                    //textBox1.AppendText($"{j},{i} created {Environment.NewLine}");
                }

                int rand = rnd.Next(0, 3);
                i += directionsX[rand];
                j += directionsY[rand];
            }
            //textBox1.AppendText($"Ended at {j},{i} {Environment.NewLine}");
        }

        public List<Vector2> GenerateNumofFloorTiles(int num)
        {
            //Make an array to store created panels
            panels = new bool[sizeX, sizeY];
            List<Vector2> CreatedPanels = new List<Vector2>();

            //Create a random starting point on the grid
            Random rnd = new Random();
            Vector2 panelCoords = new Vector2(rnd.Next(0, sizeX), rnd.Next(0, sizeY));
            //int i = ;
            //int j =;

            //only movement in the cardinal directions is allowed
            //int[] directionsX = { 0, 2, 0, -2 };
            //int[] directionsY = { 2, 0, -2, 0 };
            Vector2[] directions =
                {
                    new Vector2(0, 2),
                    new Vector2(2, 0),
                    new Vector2(0, -2),
                    new Vector2(-2,0)
                };

            //record the current coordinates as they are valid
            Vector2 lastValidCoords = panelCoords;

            while (num > 0)
            {
                //textBox1.AppendText($"{num} trying {panelCoords.ToString()} {Environment.NewLine}");

                //If there is nothing in this location create a panel
                if (panels[(int)panelCoords.X, (int)panelCoords.Y] == false)
                {

                    //Create a 2x2m panel at the selected location
                    Create2x2MPanel(panelCoords);
                    //Add the panel to the list
                    panels[(int)panelCoords.X, (int)panelCoords.Y] = true;
                    CreatedPanels.Add(panelCoords);
                    //textBox1.AppendText($"{panelCoords.ToString()} created {Environment.NewLine}");

                    //Repeat unitl num has been reached
                    num -= 1;
                }

                //Move onto the next location
                int rand = rnd.Next(0, 3);
                //i += directionsX[rand];
                //j += directionsY[rand];
                panelCoords += directions[rand];

                //Check if they are valid
                if (ContainsCoordinates(panelCoords))
                {
                    lastValidCoords = panelCoords;
                }
                else
                {
                    //textBox1.AppendText($"{num} didnt work at {panelCoords.ToString()} {Environment.NewLine}");

                    //if the coords arent valid move around until they are
                    //int k = 3;
                    List<Vector2> directionRemoval = directions.ToList<Vector2>();

                    while (!ContainsCoordinates(panelCoords) && directionRemoval.Count > 0)
                    {
                        //textBox1.AppendText($"k is:{k}{Environment.NewLine}");

                        //randomise the coordinates based on the last successful coordinate
                        //j = (int)lastValidCoords.X;
                        //i = (int)lastValidCoords.Y;
                        panelCoords = lastValidCoords;

                        int i = rnd.Next(0, directionRemoval.Count);
                        //textBox1.AppendText($"{i} next direction is {directionRemoval[i].ToString()} {Environment.NewLine}");

                        panelCoords += directionRemoval[i];

                        directionRemoval.RemoveAt(i);
                        //i += directionsX[k];
                        //j += directionsY[k];
                        //k -= 1;
                    }

                    //textBox1.AppendText($"{directionRemoval.Count} test {panelCoords.ToString()} {Environment.NewLine}");


                    //If a valid coordinate cannot be found stop running
                    if (directionRemoval.Count <= 0) {
                        return CreatedPanels;
                    }
                }

            }
            return CreatedPanels;
            //textBox1.AppendText($"Ended at {j},{i} {Environment.NewLine}");
        }

        void Create2x2MPanel(Vector2 coords)
        {
            for (int m = 0; m < 2; m++)
            {
                for (int l = 0; l < 2; l++)
                {
                    CreatePanelAtPoint(coords + new Vector2(m, l));
                }
            }
        }

        void CopyGeneratedTiles(int num)
        {
            RobotSelection robotSelection2 = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_PANEL);
            robotSelection2.FromText("all");
            robapp.Project.Structure.Edit.SelTranslate(15, 0, 0, IRobotTranslateOptions.I_TO_COPY, num);

            robotSelection2 = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_PANEL);
            robotSelection2.FromText("all");
            robapp.Project.Structure.Edit.SelTranslate(0, 15, 0, IRobotTranslateOptions.I_TO_COPY, num);
        }

        void Calculate()
        {
            IRobotStructure robotStructure = robapp.Project.Structure;
            IRDimServer rDimServer;
            RDimStream rDimStream;
            IRDimCalcEngine rDimCalcEngine;

            rDimServer = (IRDimServer)robapp.Kernel.GetExtension("RDimServer");
            rDimStream = rDimServer.Connection.GetStream();
            rDimServer.Mode = IRDimServerMode.I_DSM_STEEL;
            rDimCalcEngine = rDimServer.CalculEngine;

            RobotCalcEngine calcEngine = robapp.Project.CalcEngine;
            calcEngine.Calculate();
            rDimCalcEngine.Solve(null);
        }

        public bool ContainsCoordinates(Vector2 coords)
        {
            return coords.X >= 0 && coords.X < sizeX && coords.Y >= 0 && coords.Y < sizeY;
        }

        public void CreatePanelAtPoint(Vector2 coords)
        {
            //Create Nodes
            IRobotCollection robotNodeServer = robapp.Project.Structure.Nodes.GetAll();
            int totalnodes = robotNodeServer.Count + 1;
            int count = 0;
            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    robapp.Project.Structure.Nodes.Create(totalnodes + count, coords.X + i, coords.Y + j, 0);
                    count++;
                }
            }

            //Create Array from nodes
            RobotNumbersArray robotNumbersArray = robapp.CmpntFactory.Create(IRobotComponentType.I_CT_NUMBERS_ARRAY);
            robotNumbersArray.SetSize(4);
            int[] offsetCount = { 0, 1, 3, 2 };
            for (int i = 0; i < 4; i++)
            {
                robotNumbersArray.Set(i + 1, totalnodes + offsetCount[i]);
            }

            //Create finite element
            int totalFE = robapp.Project.Structure.FiniteElems.GetAll().Count + 1;
            robapp.Project.Structure.FiniteElems.Create(totalFE, robotNumbersArray);

            //Create panel
            int totalObjects = robapp.Project.Structure.Objects.GetAll().Count + 1;
            robapp.Project.Structure.Objects.CreateOnFiniteElems(totalFE.ToString(), totalObjects);
            IRobotObjObject panel = robapp.Project.Structure.Objects.Get(totalObjects) as IRobotObjObject;
            panel.SetLabel(IRobotLabelType.I_LT_PANEL_THICKNESS, panelSection);
            panel.SetLabel(IRobotLabelType.I_LT_PANEL_CALC_MODEL, "Shell");
        }

        public double[] GetAndDisplayNodeLocationsOnGraph()
        {
            RobotApplication robotApplication = new RobotApplication();
            //chart1.Series.Clear();

            //var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
            //{
            //    Name = "Series1",
            //    Color = System.Drawing.Color.Green,
            //    IsVisibleInLegend = false,
            //    ChartType = SeriesChartType.Point
            //};

            //this.chart1.Series.Add(series1);

            //textBox1.AppendText("Method Running" + Environment.NewLine);


            double[] nodes = new double[100];
            int count = 0;

            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {

                    nodes[count] = DoesNodeExistAtXY(new Vector2(j, i), false);
                    bool nodeReal = nodes[count] != 0;
                    if (nodeReal)
                    {
                        //series1.Points.AddXY(j, i);
                    }
                    //textBox1.AppendText($"{j},{i} node is {nodeReal.ToString()} {Environment.NewLine}");
                    count++;

                }
            }
            //chart1.Invalidate();

            return nodes;

        }

        void DeleteAll(int i)
        {

            if (i % 25 == 0)
            {
                if (PassdownChild != null)
                {
                    SaveData(PassdownChild);


                    //PressEscape();


                    //string FileName = $"{DateTime.Now.ToString()} , Gen {PassdownChild.Gen.ToString("0000")}.rtd";
                    //FileName = FileName.Replace(":", "-");
                    //FileName = FileName.Replace("/", "-");
                    //robapp.Project.SaveAs(FileName);
                    
                }



                robapp.Quit(IRobotQuitOption.I_QO_DISCARD_CHANGES);
                Process.Start(@"C:\Program Files\Autodesk\Autodesk Robot Structural Analysis Professional 2019\System\Exe\robot.EXE");
                Thread.Sleep(8000);

                robapp = new RobotApplication();
                robapp.Interactive = 0;
                robapp.UserControl = false;
                CreateMaterialsColumnsAndPanels();


            }
            robapp.Project.New(IRobotProjectType.I_PT_BUILDING);


        }

        //void PressEscape()
        //{
        //    this.WindowState = FormWindowState.Maximized;

        //    Process[] processes = Process.GetProcesses();

        //    foreach (Process proc in processes)
        //    {
        //        textBox1.AppendText(proc.MainWindowTitle + Environment.NewLine);
        //        PostMessage(proc.MainWindowHandle, WM_KEYDOWN, VK_F5, 0);
        //    }

            
            
        //} 
    

        public void UpdateMeshesForCorrectCalculationSettings()
        {

            try
            {
                RobotProjectPreferences ProjectPrefs = robapp.Project.Preferences;
                RobotMeshParams MeshParams = robapp.Project.Preferences.MeshParams;
                MeshParams.SurfaceParams.Method.Method = IRobotMeshMethodType.I_MMT_COONS;
                MeshParams.SurfaceParams.Method.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;
                MeshParams.SurfaceParams.Generation.Type = IRobotMeshGenerationType.I_MGT_ELEMENT_SIZE;
                MeshParams.SurfaceParams.Generation.ElementSize = 1;
                MeshParams.SurfaceParams.Coons.PanelDivisionType = IRobotMeshPanelDivType.I_MPDT_SQUARE_IN_RECT;
                MeshParams.SurfaceParams.Coons.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;
                MeshParams.SurfaceParams.FiniteElems.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;

                RobotSelection robotSelection = robapp.Project.Structure.Selections.Create(IRobotObjectType.I_OT_PANEL);
                robotSelection.FromText("all");

                RobotObjObjectCollection panelCol = robapp.Project.Structure.Objects.GetMany(robotSelection) as RobotObjObjectCollection;

                for (int i = 1; i <= panelCol.Count; i++)
                {
                    RobotObjObject panel = panelCol.Get(i);
                    //textBox1.AppendText($"{panel.Number.ToString()}{Environment.NewLine}");
                    panel.Main.Attribs.Meshed = 1;
                    panel.Mesh.Params.MeshType = IRobotMeshType.I_MT_USER;
                    panel.Mesh.Params.SurfaceParams.Method.Method = IRobotMeshMethodType.I_MMT_COONS;
                    panel.Mesh.Params.SurfaceParams.Method.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;
                    panel.Mesh.Params.SurfaceParams.Generation.Type = IRobotMeshGenerationType.I_MGT_ELEMENT_SIZE;
                    panel.Mesh.Params.SurfaceParams.Generation.ElementSize = 1;
                    panel.Mesh.Params.SurfaceParams.Coons.PanelDivisionType = IRobotMeshPanelDivType.I_MPDT_SQUARE_IN_RECT;
                    panel.Mesh.Params.SurfaceParams.Coons.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;
                    panel.Mesh.Params.SurfaceParams.FiniteElems.ForcingRatio = IRobotMeshForcingRatio.I_MFR_FORCED;
                    panel.Update();
                    panel.Mesh.Generate();
                }

            }
            catch (Exception e)
            {
                textBox1.AppendText(Environment.NewLine + e.ToString());
            }
            finally
            {

            }

        }

        Child ConstructNN(int id, double[] inputs, Vector2 offset)
        {

            Child NN = new Child(id, offset, 1);
            Random rnd = new Random();
            double val = 0;

            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        NN.weights[i,j,k] = ((rnd.NextDouble() * 2.0) - 1.0) * 0.00001;
                    }
                    NN.bias[i,j] = (rnd.NextDouble() * 2.0) - 1.0;
                }
            }


            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        NN.values[i,j,k] = inputs[k] * NN.weights[i, j, k];
                        val += NN.values[i, j, k];
                    }

                    val += NN.bias[i, j];
                    NN.totals[i, j] = val;
                    inputs[j] = NN.totals[i, j];
                    val = 0;
                }
            }

            int count = 0;
            List<Vector2> coords = new List<Vector2>();
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    if (inputs[count] > 0)
                    {
                        coords.Add(new Vector2(j, i));
                    }
                    count++;
                }
            }



            foreach (var item in coords)
            {
                NN.NumberOfColumns += CreateColumn(item + offset);
            }

            return NN;
        }

        Child ConstructNN(int id, double[] inputs, Vector2 offset, Child previousBest)
        {
            Child NN = new Child(id, offset, previousBest.Gen);
            int gen = previousBest.Gen;
            if (gen > 100)
            {
                gen = 100;
            }
            Random rnd = new Random();
            double val = 0;

            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        NN.weights[i, j, k] = previousBest.weights[i, j, k] + ((rnd.NextDouble() * 2.0) - 1.0) * 0.001*(1/gen);
                    }
                    //textBox1.AppendText($"weights:{i}{j}{1} is {weights[i, j, 1]}{Environment.NewLine}");

                    NN.bias[i, j] = previousBest.bias[i, j] + ((rnd.NextDouble() * 2.0) - 1.0) * 1 * (1 /gen);
                    //textBox1.AppendText($"bias:{i}{j} is {bias[i, j]}{Environment.NewLine}");

                }
            }


            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        NN.values[i, j, k] = inputs[k] * NN.weights[i, j, k];
                        val += NN.values[i, j, k];
                    }
                    //textBox1.AppendText($"value:{i}{j}{1} is {values[i, j, 1]}{Environment.NewLine}");

                    val += NN.bias[i, j];
                    NN.totals[i, j] = val;
                    //textBox1.AppendText($"total:{i}{j} is {totals[i, j]}{Environment.NewLine}");
                    inputs[j] = NN.totals[i, j];
                    val = 0;
                }
            }

            //var series2 = new System.Windows.Forms.DataVisualization.Charting.Series
            //{
            //    Name = "Series2",
            //    Color = System.Drawing.Color.Red,
            //    IsVisibleInLegend = false,
            //    ChartType = SeriesChartType.Point
            //};

            //this.chart1.Series.Add(series2);

            int count = 0;
            List<Vector2> coords = new List<Vector2>();
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    if (inputs[count] > 0)
                    {
                        //series2.Points.AddXY(j, i);
                        coords.Add(new Vector2(j, i));
                    }
                    count++;
                }
            }

            //chart1.Invalidate();

            foreach (var item in coords)
            {
                NN.NumberOfColumns += CreateColumn(item + offset);
            }
            return NN;
        }


        private void SetTopNode()
        {
            lastNode = robapp.Project.Structure.Nodes.GetAll().Count;
            lastObject = robapp.Project.Structure.Objects.GetAll().Count;
        }
        public int CreateColumn(Vector2 coords)
        {

            //textBox1.AppendText($"Start: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");
            int topnode = 0;
            int bottomnode = 0;

            //Does the top node exist?
            if (NodeLocationsAndNumbersDict.ContainsKey(coords))
            {
                NodeLocationsAndNumbersDict.TryGetValue(coords, out topnode);
            }
            else
            {
                topnode = -1;
            }

            //textBox1.AppendText($"Top Node: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

            if (topnode >= 0)
            {
                //the top node exists and is probably a low number so use the last highest node number
                bottomnode = lastNode + 1;
            }
            else
            {
                //If topnode is 0 then it doesnt exist
                //textBox1.AppendText($"Return: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                return 2;
            }
            //textBox1.AppendText($"Check bottom node: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

            while (AllNodesInModel.Contains(bottomnode))
            {
                bottomnode += 1;
            }
            //textBox1.AppendText($"BottomNode: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

            lastNode = Math.Max(lastNode, bottomnode);


            robapp.Project.Structure.Nodes.Create(topnode, coords.X, coords.Y, 0);
            robapp.Project.Structure.Nodes.Create(bottomnode, coords.X, coords.Y, -3);
            robapp.Project.Structure.Nodes.Get(bottomnode).SetLabel(IRobotLabelType.I_LT_SUPPORT, "Fixed");
            //textBox1.AppendText($"Create nodes: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");


            int totalBars = lastNode + 1;
            robapp.Project.Structure.Bars.Create(totalBars, topnode, bottomnode);
            robapp.Project.Structure.Bars.Get(totalBars).SetLabel(IRobotLabelType.I_LT_BAR_SECTION, columnSection);
            //textBox1.AppendText($"Create bar: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

            return 1;
        }


        //public int CreateColumn(Vector2 coords)
        //{
        //    int topnode = 0;
        //    int bottomnode = 0;

        //    textBox1.AppendText($"Create Columns: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");


        //    //Does the top node exist?
        //    topnode = DoesNodeExistAtXY(coords, true);
        //    //If 0 is returned it doesnt
        //    textBox1.AppendText($"Does node exist: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

        //    if (topnode == 0)
        //    {
        //        return 2;
        //        //Find a node which works
        //        topnode = lastNode + 1;
        //        while (CheckNodeExists(topnode))
        //        {
        //            //Check until this returns there is no node at that number
        //            topnode += 1;
        //        }
        //        lastNode = Math.Max(lastNode, topnode);
        //        bottomnode = topnode + 1;
        //    }
        //    else
        //    {
        //        //the top node exists and is probably a low number so use the last highest node number
        //        bottomnode = lastNode + 1;
        //    }


        //    textBox1.AppendText($"Check node exist: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

        //    while (CheckNodeExists(bottomnode))
        //    {
        //        bottomnode += 1;
        //        lastNode = Math.Max(lastNode, bottomnode);
        //    }
        //    textBox1.AppendText($"Check Node Exists: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");


        //    robapp.Project.Structure.Nodes.Create(topnode, coords.X, coords.Y, 0);
        //    robapp.Project.Structure.Nodes.Create(bottomnode, coords.X, coords.Y, -3);
        //    var support = robapp.Project.Structure.Nodes.Get(bottomnode);
        //    support.SetLabel(IRobotLabelType.I_LT_SUPPORT, "Fixed");

        //    textBox1.AppendText($"Create Supports: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");


        //    var robotBarServer = robapp.Project.Structure.Objects.GetAll();
        //    int totalBars = robotBarServer.Count + 1;
        //    robapp.Project.Structure.Bars.Create(totalBars, topnode, bottomnode);
        //    robapp.Project.Structure.Bars.Get(totalBars).SetLabel(IRobotLabelType.I_LT_BAR_SECTION, "col1");
        //    robapp.Project.Structure.Bars.Get(totalBars).SetLabel(IRobotLabelType.I_LT_MEMBER_TYPE, "Column");
        //    textBox1.AppendText($"Create column: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

        //    return 1;
        //}

        void CreateColumnSize()
        {

            var robotLabel = robapp.Project.Structure.Labels.Create(IRobotLabelType.I_LT_BAR_SECTION, "col1");
            RobotBarSectionData robotBarSectionData = robotLabel.Data;
            robotBarSectionData.ShapeType = IRobotBarSectionShapeType.I_BSST_CONCR_BEAM_RECT;
            robotBarSectionData.MaterialName = "C25";
            robotBarSectionData.Concrete.SetValue(IRobotBarSectionConcreteDataValue.I_BSCDV_BEAM_T_B, 0.25);
            robotBarSectionData.Concrete.SetValue(IRobotBarSectionConcreteDataValue.I_BSCDV_BEAM_T_H, 0.25);

            robotBarSectionData.SetValue(IRobotBarSectionDataValue.I_BSDV_BF, 1);
            robotBarSectionData.SetValue(IRobotBarSectionDataValue.I_BSDV_D, 1);
            robotBarSectionData.CalcNonstdGeometry();
            robapp.Project.Structure.Labels.Store(robotLabel);
        }

        HashSet<int> GetAllNodes()
        {
            var allNodes = new HashSet<int>();
            var numberOfNodes = robapp.Project.Structure.Nodes.GetAll().Count;

            for (int i = 0; i < numberOfNodes; i++)
            {
                var j = CheckNodeExistsReturnInt(i);
                if(j != -1)
                {
                    allNodes.Add(i);
                }
            }

            return allNodes;
        }

        int CheckNodeExistsReturnInt(int i)
        {
            try
            {
                IRobotDataObject node = robapp.Project.Structure.Nodes.Get(i);
                node.HasLabel(IRobotLabelType.I_LT_SUPPORT);
                return i;
            }
            catch (Exception e)
            {

            }
            finally
            {

            }
            return -1;
        }

        bool CheckNodeExists(int i)
        {
            try
            { 
                IRobotDataObject node = robapp.Project.Structure.Nodes.Get(i);
                node.HasLabel(IRobotLabelType.I_LT_SUPPORT);
                return true;
            }
            catch (Exception e)
            {

            }
            finally
            {
              
            }
            return false;
        }

        void PrintErrorMessage(Exception E)
        {
            textBox2.AppendText(Environment.NewLine + E.ToString());
        }

        Dictionary<Vector2, int> NodeLocationsAndNumbers(List<Vector2> coords, int tiles)
        {
            var nodeLocationsAndNumbers = new Dictionary<Vector2, int>();

            RobotStructureCache robotStructureCache = robapp.Project.Structure.CreateCache();
            foreach (var coord in coords)
            {
                for (int i = 0; i < tiles; i++)
                {
                    for (int j = 0; j < tiles; j++)
                    {
                        for (int k = 0; k < 3; k++)
                        {
                            for (int l = 0; l < 3; l++)
                            {
                                var nodeLocation = new Vector2(coord.X + 15 * i + k, coord.Y + 15 * j + l);
                                int nodeNum = robotStructureCache.EnsureNodeExist(nodeLocation.X, nodeLocation.Y, 0);
                                if (nodeNum > -1)
                                {
                                    if (nodeLocationsAndNumbers.ContainsKey(nodeLocation))
                                    {
                                        break;
                                    }
                                    nodeLocationsAndNumbers.Add(nodeLocation, nodeNum);
                                }
                            }
                        }
                    }
                }
                
            }

            return nodeLocationsAndNumbers;
        }

        public int DoesNodeExistAtXY(Vector2 coords, bool returnNodeNumber)
        {

            //What is this doing?

            //Find every call of robapp and remove it where possible
            RobotStructureCache robotStructureCache = robapp.Project.Structure.CreateCache();
            int i = robotStructureCache.EnsureNodeExist(coords.X, coords.Y, 0);
            try
            {
                IRobotDataObject node = robapp.Project.Structure.Nodes.Get(i);
                if(node.Number == i && !returnNodeNumber)
                {
                    return 1;
                }
                else
                {
                    return node.Number;
                }
            }
            catch(Exception e)
            {

            }
            finally
            {
                
            }
            return -1;
        }


        #region old code
        /*
       
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
        */
    }
#endregion

}
