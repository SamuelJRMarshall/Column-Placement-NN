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
//using Microsoft.Office.Interop.Excel;
//using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;
using static Tensorflow.Binding;
using System.Numerics;
using System.IO;

namespace Column_Sort
{
    // made by Sam
    public partial class Form1 : Form
    {
        public RobotApplication robapp;
        private bool[,] panels;
        private int sizeX = 8 , sizeY = 8;
        private int numOfColumns;
        public double DeflectionThreshold = -10;
        public double ColumnFactor = 20;

        int lastNode = 0;

        private Child PassdownChild;
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
                robapp.Interactive = 0;
                robapp.UserControl = false;




                //DeleteAll();
                //GenerateNumofFloorTiles(15);
                //UpdateMeshesForCorrectCalculationSettings();
                //Calculate();
                //CreateColumnSize();
                //GetAndDisplayNodeLocationsOnGraph();
                ////CreateColumn(4, 4);
                //Calculate();
                //GetResults();

                //Refresh Model
                DateTime startTime = DateTime.Now;
                textBox1.AppendText($"Start: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                DeleteAll();
                textBox1.AppendText($"Delete: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                //Create a floor plan
                List<Vector2> nodeLocations = GenerateNumofFloorTiles(10);
                textBox1.AppendText($"Generate Tiles: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                //Copy the plan
                int tiles = 5;

                CopyGeneratedTiles(tiles - 1);
                textBox1.AppendText($"Copy Tiles: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                //Analyse the samples
                double[] inputs = GetAndDisplayNodeLocationsOnGraph();
                textBox1.AppendText($"Get nodes: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                SetTopNode();
                textBox1.AppendText($"Set Top node: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                //Create NN
                Child[] children = new Child[tiles * tiles];
                int k = 0;
                for (int i = 0; i < tiles; i++)
                {
                    for (int j = 0; j < tiles; j++)
                    {
                        if (PassdownChild == null)
                        {
                            children[k] = new Child(k, new Vector2(15 * i, 15 * j), true, this);
                        }
                        else
                        {
                            children[k] = new Child(k, new Vector2(15 * i, 15 * j), true, this, PassdownChild.weights, PassdownChild.bias);
                        }
                        textBox1.AppendText($"Create child {k}: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                        k += 1;
                    }
                }


                //Place Columns
                foreach (var child in children)
                {
                    child.PlaceColumns(child.GenerateValuesFromInput(inputs));
                    textBox1.AppendText($"Place columsn for child {child.Id}: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                }

                Calculate();
                textBox1.AppendText($"Calculate : {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                //Get Deflections
                //Score output
                foreach (var child in children)
                {
                    child.Score = GetResults(nodeLocations, child.Offset);
                    textBox1.AppendText($"Get results for child {child.Id}: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                }

                List<Child> orderedChildren = children.OrderBy(o => o.Score).ToList();

                //foreach (var child in orderedChildren)
                //{
                //    textBox1.AppendText(child.Score.ToString() + Environment.NewLine);
                //}


                PassdownChild = orderedChildren[0];
                //textBox1.AppendText("Best: " + PassdownChild.Score.ToString() + Environment.NewLine);

                textBox1.AppendText($"Done: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");

                //save data to file
                //load data back from file
                //randomise the loaded results
                //run again
                //repeat

                //Pass Down best
                //Repeat



                //textBox1.AppendText(LoadData());
            }
            catch (Exception E)
            {
                textBox1.Text = (E.ToString());
            }
            finally
            {
                robapp.Interactive = 1;
                robapp.UserControl = true;
                robapp = null;
            }

        }

        

        double GetResults(List<Vector2> nodeLocations, Vector2 offset)
        {
            HashSet<int> existingNodes = new HashSet<int>();
            foreach (var node in nodeLocations)
            {
                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        int num = DoesNodeExistAtXY(node.X + j + offset.X, node.Y + i + offset.Y, true);
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
                else
                {
                    deflectionScore += val;
                }

                highestDeflectionZ = Math.Min(highestDeflectionZ, val);
            }
            return -(deflectionScore + highestDeflectionZ);
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
                            writetext.Write(bestChild.weights[i, j, k].ToString("G17") +"|");
                        }

                    }
                }

                writetext.WriteLine();

                for (int i = 0; i < 8; i++)
                {
                    for (int j = 0; j < 100; j++)
                    {
                        writetext.Write(bestChild.bias[i, j].ToString("G17") + "|");
                    }
                }
            }

        }

        public string LoadData()
        {
            using (StreamReader readtext = new StreamReader("data.txt"))
            {
                return readtext.ReadLine();
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

            while (ContainsCoordinates(j, i))
            {
                if (panels[j, i] == false)
                {
                    CreatePanelAtPoint(j, i);
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
            int i = rnd.Next(0, sizeX);
            int j = rnd.Next(0, sizeY);
            
            //only movement in the cardinal directions is allowed
            int[] directionsX = { 0, 2, 0, -2 };
            int[] directionsY = { 2, 0, -2, 0 };

            //record the current coordinates as they are valid
            Vector2 coords = new Vector2(j, i);

            while (num > 0)
            {
                //textBox1.AppendText($"{num} trying {j},{i} {Environment.NewLine}");

                //If there is nothing in this location create a panel
                if (panels[j, i] == false)
                {

                    //Create a 2x2m panel at the selected location
                    Create2x2MPanel(new Vector2(j,i));
                    //Add the panel to the list
                    panels[j, i] = true;
                    CreatedPanels.Add(new Vector2(j, i));
                    //textBox1.AppendText($"{j},{i} created {Environment.NewLine}");

                    //Repeat unitl num has been reached
                    num -= 1;
                }

                //Move onto the next location
                int rand = rnd.Next(0, 3);
                i += directionsX[rand];
                j += directionsY[rand];

                //Check if they are valid
                if (ContainsCoordinates(j, i))
                {
                    coords = new Vector2(j, i);
                }
                else
                {
                    //if the coords arent valid move around until they are
                    int k = 3;
                    while (!ContainsCoordinates(j, i) && k >= 0)
                    {
                        //textBox1.AppendText($"k is:{k}{Environment.NewLine}");
                        //randomise the coordinates based on the last successful coordinate
                        j = (int)coords.X;
                        i = (int)coords.Y;

                        i += directionsX[k];
                        j += directionsY[k];
                        k -= 1;
                    }

                    //If a valid coordinate cannot be found stop running
                    if(k < 0 && !ContainsCoordinates(j, i)){
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
                    CreatePanelAtPoint(coords.X + m, coords.Y + l);
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
            panel.SetLabel(IRobotLabelType.I_LT_PANEL_THICKNESS, "TH10");
            panel.SetLabel(IRobotLabelType.I_LT_MATERIAL, "C10");
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

            for (float i = 0; i < 10; i++)
            {
                for (float j = 0; j < 10; j++)
                {

                    nodes[count] = DoesNodeExistAtXY(j, i, false);
                    bool nodeReal = nodes[count] != 0;
                    if (nodeReal)
                    {
                        //series1.Points.AddXY(j, i);
                    }
                    //textBox1.AppendText($"{j},{i} node is {nodeReal.ToString()} {Environment.NewLine}");
                    count++;

                }
            }
            chart1.Invalidate();

            return nodes;
            
        }

        void DeleteAll()
        {
            //RobotSelection robotSelection = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_PANEL);
            //robotSelection.FromText("all");
            //robapp.Project.Structure.Objects.DeleteMany(robotSelection);
            //robotSelection = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_BAR);
            //robotSelection.FromText("all");
            //robapp.Project.Structure.Bars.DeleteMany(robotSelection);
            //robotSelection = robapp.Project.Structure.Selections.Get(IRobotObjectType.I_OT_NODE);
            //robotSelection.FromText("all");
            //robapp.Project.Structure.Nodes.DeleteMany(robotSelection);

            robapp.Project.New(IRobotProjectType.I_PT_BUILDING);

        }

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
                textBox1.Text = e.ToString();
            }
            finally
            {

            }

        }

        int ConstructNN(double[] inputs, Vector2 offset)
        {
            Random rnd = new Random();
            double[,,] weights = new double[8,100,100];
            double[,] bias = new double[8, 100];
            double[,,] values = new double[8, 100, 100];
            double[,] totals = new double[8, 100];
            double val = 0;

            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        weights[i,j,k] = ((rnd.NextDouble() * 2.0) - 1.0) * 0.00001;
                    }
                    //textBox1.AppendText($"weights:{i}{j}{1} is {weights[i, j, 1]}{Environment.NewLine}");

                    bias[i,j] = (rnd.NextDouble() * 2.0) - 1.0;
                    //textBox1.AppendText($"bias:{i}{j} is {bias[i, j]}{Environment.NewLine}");

                }
            }


            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        values[i,j,k] = inputs[k] * weights[i, j, k];
                        val += values[i, j, k];
                    }
                    //textBox1.AppendText($"value:{i}{j}{1} is {values[i, j, 1]}{Environment.NewLine}");

                    val += bias[i, j];
                    totals[i, j] = val;
                    //textBox1.AppendText($"total:{i}{j} is {totals[i, j]}{Environment.NewLine}");
                    inputs[j] = totals[i, j];
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
                CreateColumn(item.X + offset.X, item.Y + offset.Y);
            }
            return numOfColumns = coords.Count;
        }

        private void SetTopNode()
        {
            IRobotCollection robotNodeServer = robapp.Project.Structure.Nodes.GetAll();
            lastNode = robotNodeServer.Count;
        }
        public void CreateColumn(double x, double y)
        {
            int topnode = 0;
            int bottomnode = 0;
            DateTime startTime = DateTime.Now;

            //Does the top node exist?
            topnode = DoesNodeExistAtXY(x, y, true);
            //If 0 is returned it doesnt

            if (topnode == 0)
            {
                //Find a node which works
                topnode = lastNode + 1;
                while (CheckNodeExists(topnode))
                {
                    //Check until this returns there is no node at that number
                    topnode += 1;
                }
                lastNode = Math.Max(lastNode, topnode);
                bottomnode = topnode + 1;
            }
            else
            {
                //the top node exists and is probably a low number so use the last highest node number
                bottomnode = lastNode + 1;
            }
            //textBox1.AppendText($"CheckNodeExists Top: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}");


            while (CheckNodeExists(bottomnode))
            {
                bottomnode += 1;
            }
            //textBox1.AppendText($"CheckNodeExists Bottom: {(DateTime.Now - startTime).TotalSeconds}{Environment.NewLine}Top: {topnode}, Bottom: {bottomnode}, Last: {lastNode}{Environment.NewLine}");


            robapp.Project.Structure.Nodes.Create(topnode, x, y, 0);
            robapp.Project.Structure.Nodes.Create(bottomnode, x, y, -3);
            var support = robapp.Project.Structure.Nodes.Get(bottomnode);
            support.SetLabel(IRobotLabelType.I_LT_SUPPORT, "Base");


            var robotBarServer = robapp.Project.Structure.Objects.GetAll();
            int totalBars = robotBarServer.Count + 1;
            robapp.Project.Structure.Bars.Create(totalBars, topnode, bottomnode);
            robapp.Project.Structure.Bars.Get(totalBars).SetLabel(IRobotLabelType.I_LT_BAR_SECTION, "col1");
            robapp.Project.Structure.Bars.Get(totalBars).SetLabel(IRobotLabelType.I_LT_MEMBER_TYPE, "Column");

            //robapp.Project.Structure.Bars.Get(totalBars).SetLabel(IRobotLabelType.o, "Column");
        }

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

        public int GetNodeAtXY(double x, double y)
        {
            RobotStructureCache robotStructureCache = robapp.Project.Structure.CreateCache();
            return robotStructureCache.EnsureNodeExist(x, y, 0);
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        public int DoesNodeExistAtXY(double x, double y, bool returnNodeNumber)
        {
            RobotStructureCache robotStructureCache = robapp.Project.Structure.CreateCache();
            int i = robotStructureCache.EnsureNodeExist(x, y, 0);
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
            return 0;
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
