using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Project3
{
    public partial class Form3 : Form
    {
        string filepath = null;
        int k = 9;
        public Form3()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }



        private void button1_Click(object sender, EventArgs e)
        {
            k = Convert.ToInt32(textBox2.Text);
            
                if (textBox1.Text == "Normal")
            {
                DialogResult result = openFileDialog1.ShowDialog();

                if (result == DialogResult.OK) // Test result.
                {
                    filepath = openFileDialog1.FileName;
                }
                int cluster_count = 0;
                int rows = 0;
                int r_rows = 1;
                int t_rows = 1;
                int cols = 0;

                Microsoft.Office.Interop.Excel.Application IExcel = new Microsoft.Office.Interop.Excel.Application();
                ///IExcel.Visible = true;


                //string fileName = "C:\\Users\\Avijeet\\Desktop\\Data Mining 601\\Project 2\\cho.xlsx";
                string fileName = filepath;
                //open the workbook
                Workbook workbook = IExcel.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                //find the used range in worksheet
                Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);
                rows = worksheet.UsedRange.Rows.Count;
                cols = worksheet.UsedRange.Columns.Count;
                double accuracy_f = 0, precision_f = 0, recall_f = 0, f_measure_f = 0;

                int k_fold = 10;

                int rows_counter = 0;
                int whole_counter = 1;
                //            Random r = new Random();
                while (rows_counter < k_fold && whole_counter < rows)
                {
                    string ind = ",";

                    int t = rows / k_fold;
                    int l = 0;
                    for (int i = 1; i <= t + 1 && whole_counter <= rows; i++)
                    {

                        l++;
                        ind = ind + "," + whole_counter + ",";
                        whole_counter++;
                    }
                    rows_counter++;
                    r_rows = 1;
                    t_rows = 1;
                    int[] cl = new int[rows - l+1];
                    int[] test_cl = new int[t + 2];
                    double[,] fv = new double[rows - l+1, cols + 1];
                    double[,] test_fv = new double[t + 2, cols + 1];
                    string col_string = ",";
                    int[] ent = new int[cols];
                    Dictionary<string, int> val_string = new Dictionary<string, int>();
                    Dictionary<int, int> col_string_dic = new Dictionary<int, int>();

                    for (int row = 1; row <= rows; ++row)
                    {
                        if (!ind.Contains("," + row + ","))
                        {
                            for (int col = 1; col <= cols - 1; ++col)
                            {
                                //access each cell
                                if (valueArray[row, col] is string)
                                {
                                    string key = col + "_" + Convert.ToString(valueArray[row, col]);
                                    int entry;
                                    if (val_string.TryGetValue(key, out entry))
                                    {
                                        fv[r_rows, col] = entry;
                                    }
                                    else
                                    {
                                        ent[col]++;
                                        fv[r_rows, col] = ent[col];
                                        val_string.Add(key, ent[col]);
                                        if (!(col_string_dic.TryGetValue(col, out entry)))
                                        {
                                            col_string = col_string + "," + col + ",";
                                            col_string_dic.Add(col, 1);
                                        }


                                    }


                                }
                                else
                                {
                                    fv[r_rows, col] = Convert.ToDouble(valueArray[row, col]);
                                }

                            }
                            cl[r_rows] = Convert.ToInt32(valueArray[row, cols]);
                            r_rows++;
                        }
                        else
                        {
                            for (int col = 1; col <= cols - 1; ++col)
                            {
                                //access each cell
                                if (valueArray[row, col] is string)
                                {
                                    string key = col + "_" + Convert.ToString(valueArray[row, col]);
                                    int entry;
                                    if (val_string.TryGetValue(key, out entry))
                                    {
                                        test_fv[t_rows, col] = entry;
                                    }
                                    else
                                    {
                                        ent[col]++;
                                        test_fv[t_rows, col] = ent[col];
                                        val_string.Add(key, ent[col]);
                                        if (!(col_string_dic.TryGetValue(col, out entry)))
                                        {
                                            col_string = col_string + "," + col + ",";
                                            col_string_dic.Add(col, 1);
                                        }
                                    }

                                }
                                else
                                {
                                    test_fv[t_rows, col] = Convert.ToDouble(valueArray[row, col]);
                                }

                            }
                            test_cl[t_rows] = Convert.ToInt32(valueArray[row, cols]);
                            t_rows++;
                        }

                    }



                    //clean up stuffs
                    
                    double[,] t10 = new double[10, 3];
                    int[] assign = new int[t_rows];
                    for (int i = 1; i < t_rows; i++)
                    {
                        Dictionary<string, double> dm = new Dictionary<string, double>();
                        for (int j = 1; j < r_rows; j++)
                        {

                            double sum = 0;
                            int entry;
                            for (int m = 1; m <= cols; m++)
                            {
                                if (col_string_dic.TryGetValue(m, out entry))//// continues attr codn
                                {
                                    if (test_fv[i, m] != fv[j, m])
                                    {
                                        sum = sum + 1;
                                    }
                                }
                                else
                                {
                                    sum = sum + Math.Pow(Convert.ToDouble(test_fv[i, m]) - Convert.ToDouble(fv[j, m]), 2);
                                }
                            }
                            dm.Add(i + "," + j, sum);

                        }
                        var items = from pair in dm
                                    orderby pair.Value ascending
                                    select pair;
                        double wt0 = 0;
                        double wt1 = 0;
                        int count = 0;
                        foreach (KeyValuePair<string, double> pair in items)
                        {
                            string[] p = pair.Key.Split(',');
                            int cls = cl[Convert.ToInt32(p[1])];
                            double wt = 1.0 / (pair.Value * pair.Value);
                            if (cls == 0)
                            {
                                wt0 = wt0 + wt;
                            }
                            else
                            {
                                wt1 = wt1 + wt;
                            }
                            count++;
                            if (count == k)
                            {
                                break;
                            }
                        }
                        if (wt0 > wt1)
                        {
                            assign[i] = 0;
                        }
                        else
                        {
                            assign[i] = 1;
                        }


                    }
                    
                    double tp = 0, tn = 0, fp = 0, fn = 0;
                    for (int i = 1; i < t_rows; i++)
                    {
                        if (test_cl[i] == 1)
                        {
                            if (assign[i] == 1)
                            {
                                tp++;
                            }
                            else if (assign[i] == 0)
                            {
                                fn++;
                            }
                        }
                        else if (test_cl[i] == 0)
                        {
                            if (assign[i] == 1)
                            {
                                fp++;
                            }
                            else if (assign[i] == 0)
                            {
                                tn++;
                            }
                        }
                    }
                    double accuracy = 0, precision = 0, recall = 0, f_measure = 0;
                    accuracy = (tp + tn) / (tp + tn + fp + fn);
                    if (tp == 0 && fp == 0)
                    {
                        precision = 0;
                    }
                    else
                    {
                        precision = (tp) / (tp + fp);
                    }
                    if (tp == 0 && fn == 0)
                    {
                        recall = 0;
                    }
                    else
                    {
                        recall = (tp) / (tp + fn);
                    }
                    if (recall == 0 && precision == 0)
                    {
                        f_measure = 0;
                    }
                    else
                    {
                        f_measure = (2 * recall * precision) / (recall + precision);
                    }
                    accuracy_f = accuracy_f + accuracy;
                    precision_f = precision_f + precision;
                    recall_f = recall_f + recall;
                    f_measure_f = f_measure_f + f_measure;
                }



                label2.Text = label2.Text + (accuracy_f / Convert.ToDouble(rows_counter));
                label3.Text = label3.Text + (f_measure_f / Convert.ToDouble(rows_counter));
                label4.Text = label4.Text + (recall_f / Convert.ToDouble(rows_counter));
                label5.Text = label5.Text + (precision_f / Convert.ToDouble(rows_counter));

                workbook.Close(false, Type.Missing, Type.Missing);
                IExcel.Quit();


            }
            else
            {
                //if (result == DialogResult.OK) // Test result.
                //{
                //filepath = openFileDialog1.FileName;
                //}
                int cluster_count = 0;
                int rows = 0;
                int rows1 = 0;
                int r_rows = 1;
                int t_rows = 1;
                int cols = 0;
                int cols1 = 0;

                Microsoft.Office.Interop.Excel.Application IExcel = new Microsoft.Office.Interop.Excel.Application();
                ///IExcel.Visible = true;


                //string fileName = "C:\\Users\\Avijeet\\Desktop\\Data Mining 601\\Project 2\\cho.xlsx";
                string fileName = "E:\\data mining\\p3\\project3_dataset3_test.xlsx";
                string fileName1 = "E:\\data mining\\p3\\project3_dataset3_train.xlsx";
                //open the workbook
                Workbook workbook = IExcel.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                Workbook workbook1 = IExcel.Workbooks.Open(fileName1,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
                Worksheet worksheet1 = (Worksheet)workbook.Worksheets[1];
                //find the used range in worksheet
                Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);
                rows = worksheet.UsedRange.Rows.Count;
                cols = worksheet.UsedRange.Columns.Count;
                object[,] valueArray1 = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);
                rows1 = worksheet1.UsedRange.Rows.Count;
                cols1 = worksheet1.UsedRange.Columns.Count;

                int[] cl = new int[rows];
                int[] test_cl = new int[rows1];
                double[,] fv = new double[rows + 1, cols + 1];
                double[,] test_fv = new double[rows1 + 1, cols1 + 1];
                string col_string = ",";
                Dictionary<string, int> val_string = new Dictionary<string, int>();
                Dictionary<int, int> col_string_dic = new Dictionary<int, int>();
                int[] entry = new int[cols];
                for (int row = 1; row <= rows - 1; ++row)
                {

                    for (int col = 1; col <= cols - 1; ++col)
                    {
                        //access each cell
                        if (valueArray[row, col] is string)
                        {
                            string key = col + "_" + Convert.ToString(valueArray[row, col]);
                            int ent;
                            if (val_string.TryGetValue(key, out ent))
                            {
                                fv[r_rows, col] = ent;
                            }
                            else
                            {
                                entry[col]++;
                                fv[r_rows, col] = entry[col];
                                val_string.Add(key, entry[col]);
                                if (!(col_string_dic.TryGetValue(col, out ent)))
                                {
                                    col_string = col_string + "," + col + ",";
                                    col_string_dic.Add(col, 1);
                                }


                            }


                        }
                        else
                        {
                            fv[r_rows, col] = Convert.ToDouble(valueArray[row, col]);
                        }

                    }
                    cl[r_rows] = Convert.ToInt32(valueArray[row, cols]);
                    r_rows++;


                }
                for (int row = 1; row <= rows1 - 1; ++row)
                {
                    for (int col = 1; col <= cols1 - 1; ++col)
                    {
                        //access each cell
                        if (valueArray1[row, col] is string)
                        {
                            string key = col + "_" + Convert.ToString(valueArray1[row, col]);
                            int ent;
                            if (val_string.TryGetValue(key, out ent))
                            {
                                test_fv[t_rows, col] = ent;
                            }
                            else
                            {
                                entry[col]++;
                                test_fv[t_rows, col] = entry[col];
                                val_string.Add(key, entry[col]);
                                if (!(col_string_dic.TryGetValue(col, out ent)))
                                {
                                    col_string = col_string + "," + col + ",";
                                    col_string_dic.Add(col, 1);
                                }


                            }


                        }
                        else
                        {
                            test_fv[t_rows, col] = Convert.ToDouble(valueArray1[row, col]);
                        }

                    }
                    test_cl[t_rows] = Convert.ToInt32(valueArray1[row, cols]);
                    t_rows++;
                }


                //clean up stuffs
                workbook.Close(false, Type.Missing, Type.Missing);
                workbook1.Close(false, Type.Missing, Type.Missing);
                IExcel.Quit();
                double[,] t10 = new double[10, 3];
                int[] assign = new int[t_rows];
                for (int i = 1; i < t_rows; i++)
                {
                    Dictionary<string, double> dm = new Dictionary<string, double>();
                    for (int j = 1; j < r_rows; j++)
                    {

                        double sum = 0;
                        for (int k = 1; k <= cols; k++)
                        {
                            int ent;
                            if (!(col_string_dic.TryGetValue(k, out ent)))
                            {
                                if (test_fv[i, k] != fv[j, k])
                                {
                                    sum = sum + 1;
                                }
                            }
                            else
                            {
                                sum = sum + Math.Pow(Convert.ToDouble(test_fv[i, k]) - Convert.ToDouble(fv[j, k]), 2);
                            }

                        }
                        dm.Add(i + "," + j, sum);

                    }
                    var items = from pair in dm
                                orderby pair.Value ascending
                                select pair;
                    double wt0 = 0;
                    double wt1 = 0;
                    int count = 0;
                    foreach (KeyValuePair<string, double> pair in items)
                    {
                        string[] p = pair.Key.Split(',');
                        int cls = cl[Convert.ToInt32(p[1])];
                        double wt = 1.0 / (pair.Value * pair.Value);
                        if (cls == 0)
                        {
                            wt0 = wt0 + wt;
                        }
                        else
                        {
                            wt1 = wt1 + wt;
                        }
                        count++;
                        if (count == k)
                        {
                            break;
                        }
                    }
                    if (wt0 > wt1)
                    {
                        assign[i] = 0;
                    }
                    else
                    {
                        assign[i] = 1;
                    }


                }
                cols++;
                double tp = 0, tn = 0, fp = 0, fn = 0;
                for (int i = 1; i < t_rows; i++)
                {
                    if (test_cl[i] == 1)
                    {
                        if (assign[i] == 1)
                        {
                            tp++;
                        }
                        else if (assign[i] == 0)
                        {
                            fn++;
                        }
                    }
                    else if (test_cl[i] == 0)
                    {
                        if (assign[i] == 1)
                        {
                            fp++;
                        }
                        else if (assign[i] == 0)
                        {
                            tn++;
                        }
                    }
                }
                double accuracy = 0, precision = 0, recall = 0, f_measure = 0;
                accuracy = (tp + tn) / (tp + tn + fp + fn);
                precision = (tp) / (tp + fp);
                recall = (tp) / (tp + fn);
                f_measure = (2 * recall * precision) / (recall + precision);
                label2.Text = label2.Text + accuracy;
                label3.Text = label3.Text + f_measure;
                label4.Text = label4.Text + recall;
                label5.Text = label5.Text + precision;

            }

            //KNN(cl);
        }
    }
}