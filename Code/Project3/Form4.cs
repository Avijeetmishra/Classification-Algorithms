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
using Project3;
/// <summary>
/// Naive Bayes
/// </summary>
namespace Project3
{
    public partial class Form4 : Form
    {
        string filepath = null;
        int rows = 0;
        int cols = 0;
        int r_rows = 1;
        int t_rows = 1;
        int itr = 0;
        int q = 2;
        double[,] express = new double[100, 100];

        public Form4()
        {
            InitializeComponent();
        }
        public double Descriptor_Prior_Probability(double[,] fv, int[] cl, double[,] test_fv, int r, int cls, Dictionary<string, int> val_string, Dictionary<int, int> col_string_dic)
        {
            double[] count = new double[cols];
            for (int p = 1; p <= r_rows - 1; ++p)
            {
                for (int col = 1; col <= cols - 1; ++col)
                {
                    if (test_fv[r, col] == fv[p, col])
                    {
                        count[col]++;
                    }
                }
            }
            double sum = 1;
            for (int col = 1; col <= cols - 1; ++col)
            {
                if((count[col] / Convert.ToDouble(r_rows - 1)) != 0)
                {
                    sum = sum * (count[col] / Convert.ToDouble(r_rows - 1));
                }
                
            }
            return sum;
        }
        public double Descriptor_Posterior_Probability(double[,] fv, int[] cl, double[,] test_fv, int r, int cls, Dictionary<string, int> val_string, Dictionary<int, int> col_string_dic, double[,] gauss)
        {
            double prob=1;
            for (int col = 1; col <= cols - 1; ++col)
            {
                int flag0 = 0;
                int entry;
                if (!(col_string_dic.TryGetValue(col, out entry)))//// continues attr codn
                {
                    double p=1;
                    p = test_fv[r, col] - gauss[col, 1];
                    p = Math.Pow(p, 2) / 2;
                    p = p / (Math.Pow(gauss[col, 2], 2));
                    p = Math.Exp(-p);
                    p = p / ((Math.Pow((2*Math.PI), 0.5)) * gauss[col, 2]);
                    if (p == 0)
                    {
                        flag0 = 1;
                    }
                    else
                    {
                        prob = prob * p;
                    }
                    
                    p = 1;
                }
                else if (col_string_dic.TryGetValue(col, out entry))//// continues attr codn
                {
                    double p = 1;
                    string key = col + "_" + test_fv[r, col];
                    double s = 0, scls = 0; ;
                        for (int row = 1; row <= rows; ++row)
                        {
                            if (cl[row] == cls)
                            {
                                scls++;
                            if (test_fv[r, col] == fv[row, col])
                            {
                                s++;
                            }


                        }
                            
                        }
                        p = s / scls;
                        if (p == 0)
                        {
                            flag0 = 1;
                        }
                        else
                        {
                            prob = prob * p;
                        }
                        p =1;

                    
                    
                }
                if (flag0 == 1)
                {
                    string key = col + "_" + test_fv[r, col];
                    string k = ";";
                    int[] val = new int[50];
                    double v = 0;
                    foreach (string key_1 in val_string.Keys)
                    {
                        if (key_1.Contains(col + "_") && (!k.Contains(key_1)))
                        {
                            k = k + "_" + key + 1;
                            v++;

                        }
                    }

                    //prob = prob * (1 / (r_rows + v));
                    prob = prob * 0.1;
                    flag0 = 0;
                }
            }
            return prob;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                filepath = openFileDialog1.FileName;
            }



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

            int k_fold = 2;

            int rows_counter = 0;
            int whole_counter = 1;
            //            Random r = new Random();
            if (textBox1.Text == "Demo")
            {
                q = q + 2;
                int[] cl = new int[rows + 1];//(training data)classification values of data(row)
                //int[] test_cl = new int[t + 2];//(testing data)classification values of data(row)
                double[,] fv = new double[rows + 1, cols + 1];//(training data)feature values of all the tuples and their values
                                                              //double[] avg = new double[cols + 1];
                double[,] test_fv = new double[2, cols + 1];//(testing data)feature values all the tuples and their values

                t_rows = 1;
                r_rows = 1;
                string col_string = ",";
                Dictionary<string, int> val_string = new Dictionary<string, int>();
                Dictionary<int, int> col_string_dic = new Dictionary<int, int>();
                double h0 = 0;
                double h1 = 0;
                int[] ent = new int[cols];
                //get initial data from excel
                for (int row = 1; row <= rows; ++row)
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
                                    int entry1;
                                    if (!(col_string_dic.TryGetValue(col, out entry1)))
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

                        if (cl[r_rows] == 0)
                        {
                            h0++;
                        }
                        else if (cl[r_rows] == 1)
                        {
                            h1++;
                        }
                        r_rows++;
                   
                    
                    

                }
                String[] tbdata = textBox2.Text.Split(',');

                //test_fv[1,1] = tbdata[0];

                
                for (int i = 1; i < 5; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        string key = i + "_" + tbdata[j];
                        int entry;
                        if (val_string.TryGetValue(key, out entry))
                        {
                            test_fv[1, i] = entry;
                        }
                    }
                }
                int assign = -1 ;
                double[,] gauss0 = new double[cols, 3];
                double[,] gauss1 = new double[cols, 3];
                for (int col = 1; col <= cols - 1; ++col)
                {
                    int entry;
                    if (!(col_string_dic.TryGetValue(col, out entry)))//// continues attr codn
                    {
                        double s0 = 0, s1 = 0;
                        double sum0 = 0, sum1 = 0;

                        for (int row = 1; row <= rows; ++row)
                        {
                            if (cl[row] == 0)
                            {
                                sum0 = sum0 + fv[row, col];
                                s0++;
                            }
                            else if (cl[row] == 1)
                            {
                                sum1 = sum1 + fv[row, col];
                                s1++;
                            }
                        }
                        gauss0[col, 1] = sum0 / s0;
                        gauss1[col, 1] = sum1 / s1;
                        sum0 = 0;
                        sum1 = 0;
                        for (int row = 1; row <= rows; ++row)
                        {
                            if (cl[row] == 0)
                            {
                                sum0 = sum0 + ((fv[row, col] - gauss0[col, 1]) * (fv[row, col] - gauss0[col, 1]));

                            }
                            else if (cl[row] == 1)
                            {
                                sum1 = sum1 + ((fv[row, col] - gauss1[col, 1]) * (fv[row, col] - gauss1[col, 1]));
                            }
                        }
                        sum0 = sum0 / (s0 - 1);
                        sum1 = sum1 / (s1 - 1);
                        gauss0[col, 2] = Math.Pow(sum0, 0.5);
                        gauss1[col, 2] = Math.Pow(sum1, 0.5);
                    }
                }

                
                    double ph0 = h0 / (h0 + h1);
                    double ph1 = h1 / (h0 + h1);
                    double pxh0 = 0;
                    double phx0 = 0;
                    double pxh1 = 0;
                    double phx1 = 0;
                    double px = 1;


                    px = Descriptor_Prior_Probability(fv, cl, test_fv, 1, 0, val_string, col_string_dic);

                    pxh0 = Descriptor_Posterior_Probability(fv, cl, test_fv, 1, 0, val_string, col_string_dic, gauss0);
                    phx0 = (ph0 * pxh0) / px;
                    //express[i, q] = phx0;

                    pxh1 = Descriptor_Posterior_Probability(fv, cl, test_fv, 1, 1, val_string, col_string_dic, gauss1);
                    phx1 = (ph1 * pxh1) / px;
                    if (phx1 > 1)
                    {
                        //phx1 = pxh1 / 100;
                    }
                    //express[i, q+1] = phx1;
                    if (phx0 > phx1)
                    {
                        assign = 0;

                    }
                    else
                    {
                        assign = 1;

                    }
                label8.Text = "P(H0|X) : " + phx0 + " , P(H1|X) : " + phx1;
                label9.Text = "P(H0) : " + ph0 + " , P(H1) : " + ph1;
                label10.Text = "P(X) : " + px;
                label11.Text = "Classified as : " + assign ;

                /*for(int i = 1; i < t_rows; i++)
                {
                    express[i, 1] = test_cl[i];
                }*/

            }
            else
            {
                while (rows_counter < k_fold)
                {
                    string ind = ",";

                    int t = rows / k_fold;
                    for (int i = 1; i <= t + 1 && whole_counter <= rows; i++)
                    {


                        ind = ind + "," + whole_counter + ",";
                        whole_counter++;
                    }
                    rows_counter++;
                    q = q + 2;
                    int[] cl = new int[rows + 1];//(training data)classification values of data(row)
                    int[] test_cl = new int[t + 2];//(testing data)classification values of data(row)
                    double[,] fv = new double[rows + 1, cols + 1];//(training data)feature values of all the tuples and their values
                                                                  //double[] avg = new double[cols + 1];
                    double[,] test_fv = new double[t + 2, cols + 1];//(testing data)feature values all the tuples and their values

                    t_rows = 1;
                    r_rows = 1;
                    string col_string = ",";
                    Dictionary<string, int> val_string = new Dictionary<string, int>();
                    Dictionary<int, int> col_string_dic = new Dictionary<int, int>();
                    double h0 = 0;
                    double h1 = 0;
                    int[] ent = new int[cols];
                    //get initial data from excel
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
                                        int entry1;
                                        if (!(col_string_dic.TryGetValue(col, out entry1)))
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

                            if (cl[r_rows] == 0)
                            {
                                h0++;
                            }
                            else if (cl[r_rows] == 1)
                            {
                                h1++;
                            }
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
                                        int entry1;
                                        if (!(col_string_dic.TryGetValue(col, out entry1)))
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
                    int[] assign = new int[t_rows];
                    double[,] gauss0 = new double[cols, 3];
                    double[,] gauss1 = new double[cols, 3];
                    for (int col = 1; col <= cols - 1; ++col)
                    {
                        int entry;
                        if (!(col_string_dic.TryGetValue(col, out entry)))//// continues attr codn
                        {
                            double s0 = 0, s1 = 0;
                            double sum0 = 0, sum1 = 0;

                            for (int row = 1; row <= r_rows; ++row)
                            {
                                if (cl[row] == 0)
                                {
                                    sum0 = sum0 + fv[row, col];
                                    s0++;
                                }
                                else if (cl[row] == 1)
                                {
                                    sum1 = sum1 + fv[row, col];
                                    s1++;
                                }
                            }
                            gauss0[col, 1] = sum0 / s0;
                            gauss1[col, 1] = sum1 / s1;
                            sum0 = 0;
                            sum1 = 0;
                            for (int row = 1; row <= r_rows; ++row)
                            {
                                if (cl[row] == 0)
                                {
                                    sum0 = sum0 + ((fv[row, col] - gauss0[col, 1]) * (fv[row, col] - gauss0[col, 1]));

                                }
                                else if (cl[row] == 1)
                                {
                                    sum1 = sum1 + ((fv[row, col] - gauss1[col, 1]) * (fv[row, col] - gauss1[col, 1]));
                                }
                            }
                            sum0 = sum0 / (s0 - 1);
                            sum1 = sum1 / (s1 - 1);
                            gauss0[col, 2] = Math.Pow(sum0, 0.5);
                            gauss1[col, 2] = Math.Pow(sum1, 0.5);
                        }
                    }

                    for (int i = 1; i < t_rows; i++)
                    {
                        double ph0 = h0 / (h0 + h1);
                        double ph1 = h1 / (h0 + h1);
                        double pxh0 = 0;
                        double phx0 = 0;
                        double pxh1 = 0;
                        double phx1 = 0;
                        double px = 1;


                        px = Descriptor_Prior_Probability(fv, cl, test_fv, i, 0, val_string, col_string_dic);

                        pxh0 = Descriptor_Posterior_Probability(fv, cl, test_fv, i, 0, val_string, col_string_dic, gauss0);
                        phx0 = (ph0 * pxh0) / px;
                        //express[i, q] = phx0;

                        pxh1 = Descriptor_Posterior_Probability(fv, cl, test_fv, i, 1, val_string, col_string_dic, gauss1);
                        phx1 = (ph1 * pxh1) / px;
                        if (phx1 > 1)
                        {
                            //phx1 = pxh1 / 100;
                        }
                        //express[i, q+1] = phx1;
                        if (phx0 > phx1)
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
                    /*for(int i = 1; i < t_rows; i++)
                    {
                        express[i, 1] = test_cl[i];
                    }*/
                }
            }
            

            /*dataGridView1.ColumnCount = q;


            for (int i = 0; i < t_rows; i++)
            {
                var row = new DataGridViewRow();
                for (int j = 0; j < q; j++)
                {
                    row.Cells.Add(new DataGridViewTextBoxCell()
                    {
                        Value = express[i, j]
                    });
                    dataGridView1.Rows.Add(row);
                }
                
            }*/
           

            label2.Text = label2.Text + (accuracy_f / Convert.ToDouble(k_fold));
            label3.Text = label3.Text + (f_measure_f / Convert.ToDouble(k_fold));
            label4.Text = label4.Text + (recall_f / Convert.ToDouble(k_fold));
            label5.Text = label5.Text + (precision_f / Convert.ToDouble(k_fold));

            workbook.Close(false, Type.Missing, Type.Missing);
            IExcel.Quit();
        }
    }
}
