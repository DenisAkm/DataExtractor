using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TableDataExtractor
{
    public partial class Form1 : Form
    {
                
        List<string> SourceFiles = new List<string>();
        List<float> FrequencyValues = new List<float>();
        List<List<float>> AnswerData = new List<List<float>>();
        
        
        string[] Delimeter = new string[] { "\t", ",", ".", ";", " " };
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            comboBoxStile.SelectedIndex = 0;
            comboBoxSeparator.SelectedIndex = 0;
            comboBoxDelimeter.SelectedIndex = 0;
            openFileDialog1.InitialDirectory = Environment.CurrentDirectory;
            textBoxDirectory.Text = Environment.CurrentDirectory;
        }


        #region Operators and etc.
        private List<String> OpenSource()
        {
            List<String> sourceFileNames = new List<string>();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                for (int i = 0; i < openFileDialog1.FileNames.Length; i++)
                {
                    sourceFileNames.Add(openFileDialog1.FileNames[i]);
                }
            }
            return sourceFileNames;
        }

        #endregion

        #region Methods and Functions

        private void AddSource(List<String> sourceFileNames)
        {
            for (int i = 0; i < sourceFileNames.Count; i++)
            {
                listBoxSource.Items.Add(sourceFileNames[i]);
            }
        }

        private void SetParameters()
        {
            SourceFiles.Clear();
            FrequencyValues.Clear();
            AnswerData.Clear();
            for (int i = 0; i < listBoxSource.Items.Count; i++)
            {
                SourceFiles.Add(listBoxSource.Items[i].ToString());
            }

            for (int i = 0; i < listBoxValue.Items.Count; i++)
            {
                FrequencyValues.Add(Convert.ToSingle(listBoxValue.Items[i].ToString().Replace(".", ",")));
            }

        }

               
        private void ExtructingFile(string fileName)
        {
            #region SubRegion Reading

            StreamReader sr = new StreamReader(fileName);            
            int numberOfLines = 0;
            int valcol = Convert.ToInt32(numericUpDownColVal.Value) - 1;
            string[] substrings;
            string varline = sr.ReadLine();
            string delimeter = Delimeter[comboBoxDelimeter.SelectedIndex];
            while (varline != null)
            {
                numberOfLines++;                
                varline = sr.ReadLine();
            }

            sr = new StreamReader(fileName);
            float[] freqArr = new float[numberOfLines];
            float[] valArr = new float[numberOfLines];
            try
            {
                for (int i = 0; i < numberOfLines; i++)
                {
                    varline = sr.ReadLine();
                    if (!(varline == ""))
                    {
                        substrings = varline.Split(Convert.ToChar(delimeter));
                        freqArr[i] = Convert.ToSingle(substrings[0].Replace(".", ","));
                        valArr[i] = Convert.ToSingle(substrings[valcol].Replace(".", ","));
                    }
                }
            }
            catch (Exception)
            {
                int errorIndex = AnswerData.Count;
                MessageBox.Show("Unable to read file" + listBoxSource.Items[errorIndex].ToString() + ".", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                throw;
            }
            
            sr.Close();
            #endregion
                        
            List<float> answerValues = new List<float>(FrequencyValues.Count);
            for (int i = 0; i < FrequencyValues.Count; i++)
            {
                int pickIndex = FindElementIndex(freqArr, FrequencyValues[i]);                
                answerValues.Add(valArr[pickIndex]);
            }
            AnswerData.Add(answerValues);
        }

        private int FindElementIndex(float[] arr, float val)
        {
            
            float delta = Math.Abs(arr[0] - val);
            int index = 0;

            for (int i = 1; i < arr.Length; i++)
            {
                if (Math.Abs(arr[i]-val) < delta)
                {
                    delta = Math.Abs(arr[i] - val);
                    index = i; 
                }                
            }

            return index;
        }
        private void WritteAnswer()
        {
            String writingAdress = textBoxDirectory.Text + "\\" + textBoxTableName.Text + ".txt";
            
            int extN = 1;
            while (File.Exists(writingAdress))
            {
                writingAdress = textBoxDirectory.Text + "\\" + textBoxTableName.Text + " (" + extN + ")" + ".txt";                
                extN++;
            }

            int decplaces = Convert.ToInt32(numericUpDownDecPlaces.Value);
            string fA = "{0:0";
            string fB = "}";

            string format = fA;
            if (decplaces > 0)
            {
                format += ".";
                for (int q = 0; q < decplaces; q++)
                {
                    format += "0";
                }
            }
            format += fB;



            StreamWriter sw = new StreamWriter(writingAdress);

            
            double var;
            string vars;
            if (comboBoxStile.SelectedIndex == 0)
            {                
                for (int j = 0; j < SourceFiles.Count; j++)
                {
                    string line = "";
                    for (int i = 0; i < FrequencyValues.Count; i++)
                    {
                        var = AnswerData[j][i];
                        vars = String.Format(format, var);                        
                        if (i > 0)
                        {
                            line += "\t";
                        }
                        line += vars;
                    }
                    if (comboBoxSeparator.SelectedIndex == 1)
                    {
                        line = line.Replace(",", ".");
                    }
                    sw.WriteLine(line);    
                }                
            }
            else
            {                
                for (int j = 0; j < FrequencyValues.Count; j++)
                {
                    string line = "";              
                    for (int i = 0; i < SourceFiles.Count; i++)
                    {                       
                        var = AnswerData[i][j];
                        vars = String.Format(format, var);                        
                        if (i > 0)
                        {
                            line += "\t";
                        }
                        line += vars;
                    }
                    if (comboBoxSeparator.SelectedIndex == 1)
                    {
                        line = line.Replace(",", ".");
                    }
                    sw.WriteLine(line);
                }        
            }
            sw.Close();
        }
        private void SaveParameters()
        {
        }

        #endregion

        #region Events
        private void button1_Click(object sender, EventArgs e)
        {

            SetParameters();
            
            if (FrequencyValues.Count > 0)
            {
                for (int i = 0; i < SourceFiles.Count; i++)
                {
                    ExtructingFile(SourceFiles[i]);
                }

                WritteAnswer();
                SaveParameters();
            }
        }

        private void addFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddSource(OpenSource());
        }

        private void textBoxValue_KeyPress(object sender, KeyPressEventArgs e)
        {            
            if (e.KeyChar == '\r')
            {
                listBoxValue.Items.Add(textBoxValue.Text);
                textBoxValue.Text = "";
            }
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBoxValue.Items.Clear();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                listBoxValue.Items.RemoveAt(listBoxValue.SelectedIndex);
            }
            catch (Exception)
            {       
            }
            
        }

        private void deleteFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                listBoxSource.Items.RemoveAt(listBoxSource.SelectedIndex);
            }
            catch (Exception)
            {                                
            }
            
        }

        private void deleteAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBoxSource.Items.Clear();
        }

        private void buttonUp_Click(object sender, EventArgs e)
        {
            int direction = -1;
            // Checking selected item
            if (listBoxSource.SelectedItem == null || listBoxSource.SelectedIndex < 0)
                return; // No selected item - nothing to do

            // Calculate new index using move direction
            int newIndex = listBoxSource.SelectedIndex + direction;

            // Checking bounds of the range
            if (newIndex < 0 || newIndex >= listBoxSource.Items.Count)
                return; // Index out of range - nothing to do

            object selected = listBoxSource.SelectedItem;

            // Removing removable element
            listBoxSource.Items.Remove(selected);
            // Insert it in new position
            listBoxSource.Items.Insert(newIndex, selected);
            // Restore selection
            listBoxSource.SetSelected(newIndex, true);

        }

        private void buttonDown_Click(object sender, EventArgs e)
        {
            int direction = 1;
            // Checking selected item
            if (listBoxSource.SelectedItem == null || listBoxSource.SelectedIndex < 0)
                return; // No selected item - nothing to do

            // Calculate new index using move direction
            int newIndex = listBoxSource.SelectedIndex + direction;

            // Checking bounds of the range
            if (newIndex < 0 || newIndex >= listBoxSource.Items.Count)
                return; // Index out of range - nothing to do

            object selected = listBoxSource.SelectedItem;

            // Removing removable element
            listBoxSource.Items.Remove(selected);
            // Insert it in new position
            listBoxSource.Items.Insert(newIndex, selected);
            // Restore selection
            listBoxSource.SetSelected(newIndex, true);
        }

        private void buttonUpVal_Click(object sender, EventArgs e)
        {
            int direction = -1;
            // Checking selected item
            if (listBoxValue.SelectedItem == null || listBoxValue.SelectedIndex < 0)
                return; // No selected item - nothing to do

            // Calculate new index using move direction
            int newIndex = listBoxValue.SelectedIndex + direction;

            // Checking bounds of the range
            if (newIndex < 0 || newIndex >= listBoxValue.Items.Count)
                return; // Index out of range - nothing to do

            object selected = listBoxValue.SelectedItem;

            // Removing removable element
            listBoxValue.Items.Remove(selected);
            // Insert it in new position
            listBoxValue.Items.Insert(newIndex, selected);
            // Restore selection
            listBoxValue.SetSelected(newIndex, true);
        }

        private void buttonDownVal_Click(object sender, EventArgs e)
        {
            int direction = 1;
            // Checking selected item
            if (listBoxValue.SelectedItem == null || listBoxValue.SelectedIndex < 0)
                return; // No selected item - nothing to do

            // Calculate new index using move direction
            int newIndex = listBoxValue.SelectedIndex + direction;

            // Checking bounds of the range
            if (newIndex < 0 || newIndex >= listBoxValue.Items.Count)
                return; // Index out of range - nothing to do

            object selected = listBoxValue.SelectedItem;

            // Removing removable element
            listBoxValue.Items.Remove(selected);
            // Insert it in new position
            listBoxValue.Items.Insert(newIndex, selected);
            // Restore selection
            listBoxValue.SetSelected(newIndex, true);
        }
                

        private void button2_Click(object sender, EventArgs e)      // Folder Change
        {

            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog1.SelectedPath))
            {
                textBoxDirectory.Text = folderBrowserDialog1.SelectedPath;                
            }
            
        }
        private void PointPressed(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsLetterOrDigit(number) && number != 95 && number != 8 && number != 32) // цифры и символы, нижнее подчеркивание, клавиша BackSpace
            {
                e.Handled = true;
                MessageBox.Show("Invalid character");
            }

        }
        #endregion

        








    }
}
