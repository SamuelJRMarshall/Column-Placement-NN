using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace Column_Sort
{
    public class Child
    {
        public double Score;
        public readonly int Id;
        public readonly Vector2 Offset;
        public int NumberOfColumns;
        public int Gen;

        public double[,,] weights = new double[8, 100, 100];
        public double[,] bias = new double[8, 100];
        public double[,,] values = new double[8, 100, 100];
        public double[,] totals = new double[8, 100];

        Form1 controller;

        public Child(int id, Vector2 offset, int gen)
        {
            Id = id;
            Offset = offset;
            Gen = gen;
        }

        void CreateFreshNN() {
            Random rnd = new Random();
            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        weights[i, j, k] = ((rnd.NextDouble() * 2.0) - 1.0) * 0.001;
                    }
                    bias[i, j] = ((rnd.NextDouble() * 2.0) - 1.0) * 0.0001;
                }
            
            } 
        }

        void LoadNN(double[,,] loadedWeights, double[,] loadedBias)
        {
            Random rnd = new Random();
            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        weights[i, j, k] = loadedWeights[i, j, k] + ((rnd.NextDouble() * 2.0) - 1.0) * 0.000001;
                    }
                    bias[i, j] = loadedBias[i,j] + ((rnd.NextDouble() * 2.0) - 1.0) * 0.00001;
                }

            }
        }

        public List<Vector2> GenerateValuesFromInput(double[] inputs) {
            double val = 0;
            for (int i = 0; i < 8; i++)
            {
                for (int j = 0; j < 100; j++)
                {
                    for (int k = 0; k < 100; k++)
                    {
                        values[i, j, k] = inputs[k] * weights[i, j, k];
                        val += values[i, j, k];
                    }

                    val += bias[i, j];
                    totals[i, j] = val;
                    inputs[j] = totals[i, j];
                    val = 0;
                }
            }

            int count = 0;
            List<Vector2> coordinatesOfColumns = new List<Vector2>();
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    if (inputs[count] > 0)
                    {
                        coordinatesOfColumns.Add(new Vector2(j, i));
                    }
                    count++;
                }
            }
            NumberOfColumns = coordinatesOfColumns.Count;
            return coordinatesOfColumns;
        }

        public void PlaceColumns(List<Vector2> coordinatesOfColumns)
        {
            foreach (var columnCoord in coordinatesOfColumns)
            {
                controller.CreateColumn(columnCoord + Offset);
            }
        }
    }
}
