using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Direct_Intervention
{
    public partial class Form1 : Form
    {
         
        /** Initialization **/

        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

        TextWriter _writer = null;

        Workbook NewBook;
        Worksheet NewSheet;

        Workbook wb;
        Worksheet sheet;
        Range excelRange;

        Dictionary<string, List<string>> ClassDictionary = new Dictionary<string, List<string>>();

        List<string> Math = new List<string>() { "Algebra I", "Algebra II", "AP Calculus A/B", "AP Calculus B/C", "AP Statistics", "Coll Algebra", "Honors Algebra II", "Honors Geometry", "Honors PreCalculus", "Intermediate Algebra", "Statistics", "Intro to HS Math" };
        List<string> CA = new List<string>() { "AP English Language and Comp", "AP English Language and Comp", "AP English Lit and Comp", "British Literature & Composition", "Coll English", "ELL English Language Arts I", "ELL English Language Arts II", "Communication Arts I Lit and Comp", "Communication Arts II Lit and Comp", "English Language Arts I Lit and Comp", "English Language Arts II Lit and Comp" };
        List<string> Science = new List<string>() { "Physics", "Biology I", "Chemistry", "Human Anatomy & Physiology", "Honors Physics", "Honors Biology", "Honors Chemistry", "AP Chemistry", "AP Biology/AP Biology Lab", "AP Physics B" };
        List<string> SS = new List<string>() { "American Citizen", " AP Government Politics US", "AP Psychology", "AP United States History", "AP World History", "Honors US History", "Military History", "Psychology", "Sociology" };

        bool bSenior;

        int NextRowIdx = 2;
        int personNum = 1;
        //int Sleeper = 0;

        public Form1()
        {

            InitializeComponent();

            if (NewBook == null)
            {
                NewBook = app.Workbooks.Add();

                if (NewSheet == null)
                {
                    NewSheet = NewBook.ActiveSheet;
                }
            }

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
        }

        /** Form Methods **/

        private void Form1_Load(object sender, EventArgs e)
        {
            // Instantiate the writer
            _writer = new TextBoxStreamWriter(txtConsole);
            // Redirect the out Console stream
            Console.SetOut(_writer);

            Console.WriteLine("--- Direct Intervention Console ---");
        }

        /** Excel Methods **/

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                Console.WriteLine("--- Excel Worksheet Sucessfully Loaded ---");
                wb = app.Workbooks.Open(file);
                sheet = (Worksheet)wb.ActiveSheet;
                excelRange = sheet.UsedRange;
            }
        }

        string excel_getValue(string cellName)
        {
            string value = string.Empty;

            try
            {
                value = sheet.get_Range(cellName).get_Value().ToString().Trim();
            }
            catch
            {
                value = "";
            }
            return value;
        }

        int excel_getIntValue(string cellName)
        {
            int value = 0;

            try
            {
                value = Convert.ToInt16(sheet.get_Range(cellName).Value);
            }
            catch
            {
                value = 100;
            }
            return value;
        }

        /** Button Click Event **/

        private void Process_Click(object sender, EventArgs e)
        {
            Console.WriteLine("--- Begin Parsing Student Grade Percentages ---");
            RemovePassingStudents();

            Console.WriteLine("--- Begin Creating Student Class Dictionary ---");
            CreateClassDictionary();

            CreateTemplate(); //Copies First Row

            Console.WriteLine("--- Begin Classroom Prioritization ---");
            ClassroomPriority();

            NewBook.Save();

            FormatSheet(); //Formats Column Width

            Console.WriteLine("--- Begin Spreadsheet Cleanup ---");
            NewSheet.UsedRange.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
            NewSheet.UsedRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            //CloseExcelWorkbook();
            Console.WriteLine("--- Direct Intervention Processing Complete! ---");
            Console.WriteLine("--- Please Click Quit To Save Document ---");
        }

        public void CreateTemplate()
        {
            Random rnd = new Random();

            Range InitialCopy = sheet.get_Range("A1", "J1");
            Range InitialPaste = NewSheet.get_Range("A1", "A1");

            InitialCopy.Copy(Type.Missing);
            InitialPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll,
                Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            NewSheet.get_Range("A1", "J1").Font.Bold = true;
            NewSheet.get_Range("A1", "J1").Font.Underline = true;
            NewSheet.get_Range("A1", "J1").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            NewBook.SaveAs("Processed Direct Intervention Spreadsheet-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Year.ToString()); /*rnd.Next(10000, 99999).ToString());*/
        }

        void FormatSheet()
        {
            NewSheet.Columns.ClearFormats();
            NewSheet.Rows.ClearFormats();

            Range Used = NewSheet.UsedRange;

            Used.Columns.AutoFit();
        }

        public void RemoveRows()
        {
            Range removeRange;

            for (int i = 2; i < NewSheet.UsedRange.Rows.Count; i++)
            {
                NewSheet.Columns.ClearFormats();
                NewSheet.Rows.ClearFormats();

                string Cell = "A" + i.ToString();
                string value = excel_getValue(Cell);

                if (((Microsoft.Office.Interop.Excel.Range)NewSheet.Cells["1", i.ToString()]).Text != "") //FIXME
                {
                    Console.WriteLine("Not");
                    removeRange = NewSheet.get_Range(Cell, Type.Missing).EntireRow;
                    removeRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                    //i--; //Maybe Bad
                }
                else Console.WriteLine("Contains");
            }
        }

        void ClassroomPriority()
        {
            var KeyList = new List<string>();
            KeyList = ClassDictionary.Keys.ToList();

            foreach (string Key in KeyList) //Loops through all students failing classes
            {
                Console.WriteLine("--- Current Person Progress " + "<" + personNum + ">" + " ---");
                var ClassList = new List<string>();
                ClassDictionary.TryGetValue(Key, out ClassList);

                if (IsSenior(Key))
                {
                    bSenior = true;
                }
                else
                {
                    bSenior = false;
                }
                personNum++;
                Select(ClassList, Key);
            }
        }

        void Select(List<string> values, String key)
        {
            bool bClassSelected = false;

            if (!bSenior)
            {
                foreach (var studValue in values) //Loops through all of a SINGLE students failed classes
                {
                    if (bClassSelected == false)
                    {
                        //First MATH
                        foreach (var x in Math)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));

                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("Math Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                break; //Goes To Next Student
                            }
                        }
                    }
                    else break;
                    
                    //Second CA
                    if (bClassSelected == false)
                    {
                        foreach (var x in CA)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));

                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("CA Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                break; //Goes To Next Student
                            }
                        }
                    }
                    else break;

                    //Third Science
                    if (bClassSelected == false)
                    {
                        foreach (var x in Science)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));

                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("Science Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                break; //Goes To Next Student
                            }
                        }
                    }
                    else break;

                    //Fourth SS
                    if (bClassSelected == false)
                    {
                        foreach (var x in CA)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));

                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("SS Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                break; //Goes To Next Student
                            }
                        }
                    }
                    else break;

                    //Fith NONCORE
                    if (bClassSelected == false)
                    {
                        //Highest Grade?
                        //Console.WriteLine("Non-Core Equal!");
                        SelectClassName(studValue, key);
                        ClassDictionary.Remove(key);
                        bClassSelected = true;
                        break;
                    }
                    break;
                }
            }
            //Senior
            else if (bSenior)
            {
                foreach (var studValue in values) //Loops through all of a SINGLE students failed classes
                {
                    //First CA
                    if (bClassSelected == false)
                    {
                        foreach (var x in CA)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));
                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("CA Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                break; //Goes To Next Student
                            }
                        }
                    }
                    else break;

                    if (bClassSelected == false)
                    {
                        //Second MATH
                        foreach (var x in Math)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));
                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("Math Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                break; //Goes To Next Student
                            }
                        }
                    }
                    else break;

                    //Third Science
                    if (bClassSelected == false)
                    {
                        foreach (var x in Science)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));

                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("Science Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                return; //Goes To Next Student
                            }
                        }
                    }
                    else break;

                    //Fourth SS
                    if (bClassSelected == false)
                    {
                        foreach (var x in CA)
                        {
                            //Console.WriteLine(studValue + " : " + x + " Compare Value: " + studValue.CompareTo(x));

                            if (studValue.CompareTo(x) == 0)
                            {
                                //Console.WriteLine("SS Equal!");
                                SelectClassName(studValue, key);
                                bClassSelected = true;
                                ClassDictionary.Remove(key);
                                break; //Goes To Next Student
                            }
                        }
                    }
                    else break;

                    //Fith NONCORE
                    if (bClassSelected == false)
                    {
                        //Highest Grade?
                        //Console.WriteLine("Non-Core Equal!");
                        SelectClassName(studValue, key);
                        bClassSelected = true;
                        ClassDictionary.Remove(key);
                        break;
                    }
                    break;
                }
            }
        }      

        void SelectClassName(string name, string id)
        {
            Range CopyRange;
            Range PasteRange;

            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();


            for (int i = 2; i < sheet.UsedRange.Rows.Count; i++)
            {
                string ClassNameCell = "E" + i.ToString();
                string IdCell = "J" + i.ToString();

                if ((excel_getValue(ClassNameCell).CompareTo(name) == 0) && (excel_getValue(IdCell).CompareTo(id) == 0))
                {
                    //Console.WriteLine("Selected Correct Class!");
                    CopyRange = sheet.get_Range("A" + i.ToString()).EntireRow;
                    PasteRange = NewSheet.get_Range("A" + NextRowIdx.ToString()); //Dont Need To Remove Other Rows

                    CopyRange.Copy(Type.Missing);
                    NewSheet.Paste(PasteRange);

                    NextRowIdx++;
                    /** Clearing caused the names to not be copied and pasted **/
                    break;
                }     
            } 
        }

        bool IsSenior(string key)
        {
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            for (int i = 2; i < sheet.UsedRange.Rows.Count; i++)
            {
                string GradeCell = "J" + i.ToString();
                
                if (excel_getValue(GradeCell).CompareTo(key) == 0) //If Key equals StudentID
                {
                    string grade = excel_getValue("B" + i.ToString());

                    if (grade.CompareTo("12") == 0)
                    {
                        //Console.WriteLine("Is Senior!");
                        return true;
                    }
                    break;
                }
            }
            //Console.WriteLine("Not Senior!");
            return false;
        }
      
        void CreateClassDictionary()
        {     
            sheet = (Worksheet)wb.ActiveSheet;

            string StudentCell;
            string ClassCell;

            string StudentID;
            string ClassName;

            for (int j = 2; j < sheet.UsedRange.Rows.Count; j++)
            {
                StudentCell = "J" + j.ToString();
                ClassCell = "E" + j.ToString();
                StudentID = excel_getValue(StudentCell);
                ClassName = excel_getValue(ClassCell);

                if (!ClassDictionary.ContainsKey(StudentID))
                {
                    ClassDictionary.Add(StudentID, new List<string>());
                }

                ClassDictionary[StudentID].Add(ClassName);
            }
            Console.WriteLine("--- Completed Creating Student Class Dictionary ---");
        }

        void RemovePassingStudents()
        {
            Range removePassingRange;
            string percentCell;
            int percent;

            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            for (int i = 2; i < sheet.UsedRange.Rows.Count; i++)
            {
                sheet.Columns.ClearFormats();
                sheet.Rows.ClearFormats();

                percentCell = "D" + i.ToString();
                percent = excel_getIntValue(percentCell);

                //Console.Write("Cell: " + percentCell + " / " + sheet.UsedRange.Rows.Count);
                    
                if (percent >= 60)
                {
                    removePassingRange = sheet.get_Range("D" + i.ToString(), Type.Missing).EntireRow;
                    removePassingRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                    i--;
                }
            }
            //CloseExcelWorkbook();
            Console.WriteLine("--- Completed Parsing Student Grade Percentages ---");     
        }

        /** Quit & Closing Methods **/

        private void Quit_Click(object sender, EventArgs e)
        {
            try
            {
                //Maybe add warning popup
                NewBook.Save();
                NewBook.Close(false);
                wb.Close(false);
                //First Book
                releaseObj(sheet);
                releaseObj(wb);
                //Second Book
                releaseObj(NewSheet);
                releaseObj(NewBook);
            }
            catch
            {
                Console.WriteLine("An Error Has Occured: ");
                app.Quit();
                this.Close();
            }
            finally
            {
                //Application
                if (app != null)
                {
                    app.Quit();
                }
                releaseObj(app);
                this.Close();
            }
        }

        private void releaseObj(object obj) // note ref!
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
            obj = null;
        }
    }
}
