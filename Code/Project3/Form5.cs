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
/// <summary>
/// //////////Random Forest
/// </summary>
namespace Project3
{
    public partial class Form5 : Form
    {
        string filepath = null;
        int rows = 0;
        int cols = 0;
        int r_rows = 1;
        int r_cols = 1;
        int data_r_rows = 1;
        int t_rows = 1;
        int itr = 0;
        int bag = 9;
        public Form5()
        {
            InitializeComponent();
        }
        private double[] get_gini(int ind, int[] B, int[] cl, double[,] fv)
        {
            //GTE 0
            //LT 0
            //GTE 1
            //LT 1
            ///sort the index on basis of data
            ///int min 
            ///

            double[] ret_arr = new double[2];
            double min = 999999999;
            double split = 0;
            double[,] fv_temp = new double[B.Length, 3];

            for (int i = 1; i < B.Length; i++)
            {
                fv_temp[i, 0] = fv[B[i], ind];
                fv_temp[i, 1] = B[i];
                fv_temp[i, 2] = cl[B[i]];
            }

            System.Data.DataTable dt = new System.Data.DataTable();
            // assumes first row contains column names:
            dt.Columns.Add("RowValue", typeof(double));
            dt.Columns.Add("Column", typeof(double));
            dt.Columns.Add("Rowstate", typeof(double));
            // load data from string array to data table:
            for (int rowindex = 1; rowindex < fv_temp.GetLength(0); rowindex++)
            {
                DataRow row = dt.NewRow();
                for (int col = 0; col < fv_temp.GetLength(1); col++)
                {
                    row[col] = fv_temp[rowindex, col];
                }
                dt.Rows.Add(row);
            }
            // sort by first column:
            DataRow[] sortedrows = dt.Select("", "RowValue");
            /*
             * code for sorted rows to sort_arr
            */
            double[,] sort_arr = new double[1000, 3];
            for (int i = 1; i <= sortedrows.Length; i++)
            {
                object[] item = sortedrows[i - 1].ItemArray.ToArray();
                sort_arr[i, 0] = Convert.ToDouble(item[0]);
                sort_arr[i, 1] = Convert.ToDouble(item[1]);
                sort_arr[i, 2] = Convert.ToDouble(item[2]);

            }





            for (int i = 1; i < B.Length; i++)
            {
                //check previous value and proceed only if it is different
                if (i > 0 && sort_arr[i, 0] != sort_arr[i - 1, 0])
                {
                    double v = (sort_arr[i, 0] + sort_arr[i - 1, 0]) / 2;
                    double l0 = 0, l1 = 0, g0 = 0, g1 = 0;
                    for (int j = 1; j < B.Length; j++)
                    {
                        if (sort_arr[j, 0] <= v)
                        {
                            if (sort_arr[j, 2] == 0)
                            {
                                l0++;
                            }
                            else
                            {
                                l1++;
                            }

                        }
                        else
                        {
                            if (sort_arr[j, 2] == 0)
                            {
                                g0++;
                            }
                            else
                            {
                                g1++;
                            }
                        }
                    }
                    double ginil = 1 - ((l0 / (l0 + l1)) * (l0 / (l0 + l1))) - ((l1 / (l0 + l1)) * (l1 / (l0 + l1)));
                    double ginig = 1 - ((g0 / (g0 + g1)) * (g0 / (g0 + g1))) - ((g1 / (g0 + g1)) * (g1 / (g0 + g1)));
                    double gini = (((l0 + l1) / (l0 + l1 + g0 + g1)) * ginil) + (((g0 + g1) / (l0 + l1 + g0 + g1)) * ginig);
                    if (gini < min)
                    {
                        min = gini;
                        split = v;
                    }
                }

            }

            ret_arr[0] = min;
            ret_arr[1] = split;
            return ret_arr;
        }
        private double[] get_gini_string(int ind, int[] B, int[] cl, double[,] fv)
        {
            double[] ret_arr = new double[2];
            double min = 999999999;
            double split = 0;
            double[,] fv_temp = new double[B.Length, 3];

            for (int i = 1; i < B.Length; i++)
            {
                fv_temp[i, 0] = fv[B[i], ind];
                fv_temp[i, 1] = B[i];
                fv_temp[i, 2] = cl[B[i]];
            }


            string vstr = ";";
            for (int i = 1; i < B.Length; i++)
            {
                //check previous value and proceed only if it is different

                double v = fv_temp[i, 0];
                if (!(vstr.Contains(Convert.ToString(v))))
                {
                    vstr = vstr + ";" + Convert.ToString(v);

                    double l0 = 0, l1 = 0, g0 = 0, g1 = 0;
                    for (int j = 1; j < B.Length; j++)
                    {
                        if (fv_temp[j, 0] == fv_temp[i, 0])
                        {
                            if (fv_temp[j, 2] == 0)
                            {
                                l0++;
                            }
                            else
                            {
                                l1++;
                            }

                        }
                        else
                        {
                            if (fv_temp[j, 2] == 0)
                            {
                                g0++;
                            }
                            else
                            {
                                g1++;
                            }
                        }
                    }
                    double ginil = 1 - ((l0 / (l0 + l1)) * (l0 / (l0 + l1))) - ((l1 / (l0 + l1)) * (l1 / (l0 + l1)));
                    double ginig = 1 - ((g0 / (g0 + g1)) * (g0 / (g0 + g1))) - ((g1 / (g0 + g1)) * (g1 / (g0 + g1)));
                    double gini = (((l0 + l1) / (l0 + l1 + g0 + g1)) * ginil) + (((g0 + g1) / (l0 + l1 + g0 + g1)) * ginig);
                    if (gini < min)
                    {
                        min = gini;
                        split = v;
                    }
                }
            }

            ret_arr[0] = min;
            ret_arr[1] = split;
            return ret_arr;
        }
        private double[] best_split(int[] A, int[] cl, double[,] fv, Dictionary<int, int> col_string_dic)
        {
            double min = 999999999999;
            int col_best = 0;
            double split_value = 0;
            double[] ret_arr = new double[2];
            string ind_3 = ",";
            Random n = new Random();
            int col_count;
            for (int i = 0; i <= r_cols; i++)
            {
                col_count = n.Next(1, cols);
                if (!(ind_3.Contains(Convert.ToString(col_count))))
                {
                    ind_3 = ind_3 + "," + col_count + ",";
                }
                else
                {
                    i--;
                }
                

            }
            for (int i = 1; i < cols; i++)
            {
                if (ind_3.Contains("," + i + ","))
                {
                    int entry;
                    double[] res;

                    if (col_string_dic.TryGetValue(i, out entry))//// continues attr codn
                    {
                        res = get_gini_string(i, A, cl, fv);//column number, list of index
                    }
                    else
                    {
                        res = get_gini(i, A, cl, fv);//column number, list of index
                    }
                    if (res[0] < min)
                    {
                        min = res[0];
                        col_best = i;
                        split_value = res[1];
                    }
                }

                    //0 index GINI value
                    //1 index split value
                    
            }
            //0 column number on which split is taking place
            //1 value on which split is taking place
            ret_arr[0] = col_best;
            ret_arr[1] = split_value;
            return ret_arr;
        }

        private void partition(m_tree S, int[] cl, double[,] fv, Dictionary<int, int> col_string_dic)
        {
            double[] S1 = new double[2];
            int[] S_left = new int[S.index_name.Length];
            int[] S_right = new int[S.index_name.Length];
            int count_left = 1;
            int count_right = 1;
            itr++;
            if (S.index_name.Length == 2)
            {
                S.leaf_node = true;
                S.value = cl[S.index_name[1]];
                return;
            }
            int p = 2;
            for (int i = 2; i < S.index_name.Length; i++)
            {
                if (cl[S.index_name[i]] != (cl[S.index_name[i - 1]]))
                {
                    break;
                }
                else p++;
            }
            if (p == S.index_name.Length)
            {
                S.leaf_node = true;
                S.value = cl[S.index_name[1]];
                return;
            }
            S1 = best_split(S.index_name, cl, fv,col_string_dic);
            int entry;
            if (col_string_dic.TryGetValue(Convert.ToInt32(S1[0]), out entry))//// continues attr codn
            {
                for (int i = 1; i < S.index_name.Length; i++)
                {

                    if (fv[S.index_name[i], Convert.ToInt32(S1[0])] != S1[1])
                    {
                        S_left[count_left] = S.index_name[i];
                        count_left++;
                    }
                    else
                    {
                        S_right[count_right] = S.index_name[i];
                        count_right++;
                    }
                }
            }
            else
            {
                for (int i = 1; i < S.index_name.Length; i++)
                {

                    if (fv[S.index_name[i], Convert.ToInt32(S1[0])] <= S1[1])
                    {
                        S_left[count_left] = S.index_name[i];
                        count_left++;
                    }
                    else
                    {
                        S_right[count_right] = S.index_name[i];
                        count_right++;
                    }
                }
            }
            int[] tr_left = new int[count_left];
            int[] tr_right = new int[count_right];
            Array.Resize(ref S_left, count_left);
            Array.Resize(ref S_right, count_right);
            S.index_left = S_left;
            S.index_right = S_right;
            S.column = Convert.ToInt32(S1[0]);
            S.value = S1[1];
            m_tree s_tree_left = new m_tree(0, 0, S.index_left);
            m_tree s_tree_right = new m_tree(0, 0, S.index_right);
            S.lnode = s_tree_left;
            S.rnode = s_tree_right;
            if (S.lnode != null)
            {
                partition(S.lnode, cl, fv, col_string_dic);
            }
            if (S.lnode != null)
            {
                partition(S.rnode, cl, fv, col_string_dic);
            }
            return;
        }

        private int classification(double[,] test_fv, int i, m_tree node, Dictionary<int, int> col_string_dic)
        {
            int state = -1;
            if (node.leaf_node == true)
            {
                return Convert.ToInt32(node.value);
            }
            int entry;
            if (col_string_dic.TryGetValue(Convert.ToInt32(node.column), out entry))//// continues attr codn
            {
                if (test_fv[i, Convert.ToInt32(node.column)] != node.value)
                {
                    state = classification(test_fv, i, node.lnode, col_string_dic);
                }
                else
                {
                    state = classification(test_fv, i, node.rnode, col_string_dic);
                }
            }
            else
            {
                if (test_fv[i, Convert.ToInt32(node.column)] <= node.value)
                {
                    state = classification(test_fv, i, node.lnode, col_string_dic);
                }
                else
                {
                    state = classification(test_fv, i, node.rnode, col_string_dic);
                }
            }
            return state;
        }

        private void button1_Click(object sender, EventArgs e)
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
            r_cols = Convert.ToInt32(.2 * cols) + 1;

            double accuracy_f = 0, precision_f = 0, recall_f = 0, f_measure_f = 0;

            int k_fold = 10;

            int rows_counter = 0;
            int whole_counter = 1;
            //            Random r = new Random();
            while (rows_counter < k_fold && whole_counter < rows)
            {
                string ind = ",";
                int k = 0;
                int t = rows / k_fold;
                for (int i = 1; i <= t + 1 && whole_counter <= rows; i++)
                {
                    k++;
                    ind = ind + "," + whole_counter + ",";
                    whole_counter++;
                }
                rows_counter++;
                int[] data_cl = new int[rows + 1];//(training data)classification values of data(row)
                int[] test_cl = new int[t + 2];//(testing data)classification values of data(row)
                double[,] data_fv = new double[rows + 1, cols + 1];//(training data)feature values of all the tuples and their values
                                                                   //double[] avg = new double[cols + 1];
                double[,] test_fv = new double[t + 2, cols + 1];//(testing data)feature values all the tuples and their values

                
                t_rows = 1;
                data_r_rows = 1;
                r_rows = 1;
                string col_string = ",";
                int[] ent = new int[cols];
                Dictionary<string, int> val_string = new Dictionary<string, int>();
                Dictionary<int, int> col_string_dic = new Dictionary<int, int>();

                //get initial data from excel
                for (int row = 1; row <= rows-1; ++row)
                {
                    if (!ind.Contains("," + row + ","))
                    {
                        //tr_name[r_rows] = row;
                        for (int col = 1; col <= cols - 1; ++col)
                        {
                            //access each cell

                            if (valueArray[row, col] is string)
                            {
                                string key = col + "_" + Convert.ToString(valueArray[row, col]);
                                int entry;
                                if (val_string.TryGetValue(key, out entry))
                                {
                                    data_fv[data_r_rows, col] = entry;
                                }
                                else
                                {
                                    ent[col]++;
                                    data_fv[data_r_rows, col] = ent[col];
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
                                data_fv[data_r_rows, col] = Convert.ToDouble(valueArray[row, col]);
                            }

                        }
                        data_cl[data_r_rows] = Convert.ToInt32(valueArray[row, cols]);
                        data_r_rows++;
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
                
                int[,] assign = new int[t_rows,bag];
                int rows_counter_2 = 0;
                while (rows_counter_2 < bag)
                {
                    int t_2 = Convert.ToInt32(.632 * data_r_rows);
                    int[] ind_2 = new int[t_2+1];
                    Random r = new Random();
                    int cluster_count;
                    for (int i = 0; i <= t_2; i++)
                    {
                        cluster_count = r.Next(1, data_r_rows);
                        ind_2[i] = cluster_count;

                    }

                    int[] tr_name = new int[t_2 + 1];
                    r_rows = 1;
                    int[] cl = new int[data_r_rows + 1];//(training data)classification values of data(row)
                    double[,] fv = new double[data_r_rows + 1, cols];//(training data)feature values of all the tuples and their values
                    for (int rw = 1; rw <= t_2; ++rw)
                    {
                        tr_name[r_rows] = ind_2[rw];
                        for (int c = 1; c <= cols - 1; ++c)
                        {
                            //access each cell
                            fv[r_rows, c] = data_fv[ind_2[rw], c];

                        }
                        cl[r_rows] = data_cl[ind_2[rw]];
                        r_rows++;
                    }
                    int[] temp_init = new int[1];

                    m_tree head = new m_tree(0, 0.0, temp_init);
                    head.index_name = tr_name;

                    partition(head, cl, fv, col_string_dic);
                   
                    for (int i = 1; i < t_rows; i++)
                    {
                        assign[i, rows_counter_2] = classification(test_fv, i, head, col_string_dic);

                    }

                    rows_counter_2++;
                }
                int[] assign_final = new int[t_rows];

                for (int i = 1; i < t_rows; i++)
                {
                    int z = 0, o = 0;
                    for (int j = 0; j < bag; j++)
                    {
                        if (assign[i, j] == 0)
                        {
                            z++;
                        }
                        else if (assign[i, j] == 1)
                        {
                            o++;
                        }
                    }
                    if (z > o)
                    {
                        assign_final[i] = 0;
                    }
                    else
                    {
                        assign_final[i] = 1;
                    }

                }





                //clean up stuffs
                double tp = 0, tn = 0, fp = 0, fn = 0;
                for (int i = 1; i < t_rows; i++)
                {
                    if (test_cl[i] == 1)
                    {
                        if (assign_final[i] == 1)
                        {
                            tp++;
                        }
                        else if (assign_final[i] == 0)
                        {
                            fn++;
                        }
                    }
                    else if (test_cl[i] == 0)
                    {
                        if (assign_final[i] == 1)
                        {
                            fp++;
                        }
                        else if (assign_final[i] == 0)
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

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
    

    }

