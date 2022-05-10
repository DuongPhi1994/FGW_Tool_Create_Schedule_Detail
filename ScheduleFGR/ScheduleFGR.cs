using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;

namespace ScheduleFGR
{
    public partial class GreenWichSchedule : Form
    {
        public GreenWichSchedule()
        {
            InitializeComponent();
        }

        private void btnOpenData_Click(object sender, EventArgs e)
        {
            if(DialogResult.OK == this.openFileDialogData.ShowDialog())
            {
                string fileName = this.openFileDialogData.FileName;

                try
                {
                    LoadData(fileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Load data terminated!", "Load data error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            DateTime startDateFollow = DateTime.Now;
            foreach(ClassSchedule cs in classSchedules)
            {
                if(cs.OrdinalSubject == 0 || cs.OrdinalSubject == 1)
                {
                    cs.ScheduleDetails = GetListScheduleDetail(cs, cs.StartDate);
                    if (cs.TotalSlot == 80)
                        cs.ScheduleDetails.AddRange(GetListScheduleDetail(cs, cs.StartDatePart2));
                    if (cs.OrdinalSubject > 0)
                        // luu lai ngay tiep theo cua ngay cuoi cung mon hoc theo thu tu
                        startDateFollow = cs.ScheduleDetails[cs.ScheduleDetails.Count - 1].Date.AddDays(1);
                }
                else
                {
                    cs.ScheduleDetails = GetListScheduleDetail(cs, startDateFollow);
                    if (cs.OrdinalSubject > 0)
                        // luu lai ngay tiep theo cua ngay cuoi cung mon hoc theo thu tu
                        startDateFollow = cs.ScheduleDetails[cs.ScheduleDetails.Count - 1].Date.AddDays(1);
                }
            }
            List<CheckResult> listResult = new List<CheckResult>();
            //check class
            List<int> listDupClass = new List<int>();
            List<string> listClass = GetListClass();
            foreach(string cla in listClass)
            {
                listDupClass.AddRange(CheckClass(cla));
            }
            listDupClass.Sort();
            listDupClass.Distinct().ToList();
            string listDupClassString = "";
            foreach (int row in listDupClass)
            {
                listDupClassString += row + ", ";
            }
            if (listDupClassString.Length > 2)
                listDupClassString = listDupClassString.Substring(0, listDupClassString.Length - 2);
            listResult.Add(new CheckResult("Duplicate class at row:", listDupClassString));

            //check teacher
            List<int> listDupTeacher = new List<int>();
            List<string> listTeacher = GetListTeacher();
            foreach(string teacher in listTeacher)
            {
                listDupTeacher.AddRange(CheckTeacher(teacher));
            }
            listDupTeacher.Sort();
            listDupTeacher.Distinct().ToList();
            string listDupTeacherString = "";
            foreach (int row in listDupTeacher)
            {
                listDupTeacherString += row + ", ";
            }
            if(listDupTeacherString.Length > 2)
                listDupTeacherString = listDupTeacherString.Substring(0, listDupTeacherString.Length - 2);
            listResult.Add(new CheckResult("Duplicate teacher at row:",listDupTeacherString));

            //check room
            List<int> listDupRoom = new List<int>();
            List<string> listRoom = GetListRoom();
            foreach (string room in listRoom)
            {
                listDupRoom.AddRange(CheckRoom(room));
            }
            listDupRoom.Sort();
            listDupRoom.Distinct().ToList();
            string listDupRoomString = "";
            foreach (int row in listDupRoom)
            {
                listDupRoomString += row + ", ";
            }
            if (listDupRoomString.Length > 2)
                listDupRoomString = listDupRoomString.Substring(0, listDupRoomString.Length - 2);
            listResult.Add(new CheckResult("Duplicate room at row:", listDupRoomString));

            //display duplicate
            dgvListData.DataSource = listResult;
            this.dgvListData.Columns[0].Width = 200;//name
            this.dgvListData.Columns[0].ReadOnly = true;
            this.dgvListData.Columns[1].Width = this.dgvListData.Width - 260;
            this.dgvListData.Columns[1].ReadOnly = true;

            if (listDupClassString.Length == 0 && listDupTeacherString.Length == 0 && listDupRoomString.Length == 0)
                this.btnExportSchedule.Enabled = true;
            else
                this.btnExportSchedule.Enabled = false;
        }
        List<ClassSchedule> classSchedules = new List<ClassSchedule>();
        //List<DateTime> listDay_Off;

        List<string> GetListRoom()
        {
            List<string> listRoom = new List<string>();
            foreach(ClassSchedule cs in classSchedules)
            {
                if (!listRoom.Contains(cs.Room))
                    listRoom.Add(cs.Room);
            }
            return listRoom;
        }
        List<int> CheckClass(string cla)
        {
            List<int> listDuplicate = new List<int>();
            for (int i = 0; i < classSchedules.Count; i++)
            {
                ClassSchedule cs1 = classSchedules[i];
                if (!cs1.StudentGroup.Equals(cla))
                    continue;
                for (int j = i + 1; j < classSchedules.Count; j++)
                {
                    ClassSchedule cs2 = classSchedules[j];
                    if (!cs2.StudentGroup.Equals(cla))
                        continue;
                    if (IsSlotDateDuplicate(cs1, cs2))
                        listDuplicate.Add(j + 2);
                }
            }

            return listDuplicate;
        }

        List<int> CheckRoom(string room)
        {
            List<int> listDuplicate = new List<int>();
            for (int i = 0; i < classSchedules.Count; i++)
            {
                ClassSchedule cs1 = classSchedules[i];
                if (!cs1.Room.Equals(room))
                    continue;
                for (int j = i + 1; j < classSchedules.Count; j++)
                {
                    ClassSchedule cs2 = classSchedules[j];
                    if (!cs2.Room.Equals(room))
                        continue;
                    if (IsSlotDateDuplicate(cs1, cs2))
                        listDuplicate.Add(j + 2);
                }
            }

            return listDuplicate;
        }
        //con kha nang toi uu
        public bool IsSlotDateDuplicate(ClassSchedule cs1, ClassSchedule cs2)
        {
            foreach (ScheduleDetail sd1 in cs1.ScheduleDetails)
            {
                int indexDateEqual = BinarySearchDate(cs2.ScheduleDetails, sd1.Date, 0, cs2.ScheduleDetails.Count - 1);
                if (indexDateEqual != -1)
                {
                    if (sd1.Slot == cs2.ScheduleDetails[indexDateEqual].Slot)
                        return true;
                }
            }
            return false;
        }
        public int BinarySearchDate (List<ScheduleDetail> cd, DateTime date, int l, int r)
        {
            int mid = (l + r) / 2;
            if (r - l < 0)
                return -1;
            if (date == cd[mid].Date)
                return mid;
            else if( date > cd[mid].Date)
            {
                return BinarySearchDate(cd, date, mid + 1, r);
            }
            else
            {
                return BinarySearchDate(cd, date, l, mid - 1);
            }
        }
        /*List<int> CheckRoom(string room)
        {
            List<int> listDuplicate = new List<int>();
            int[,] schedule = { { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 }};
            bool isDuplicate = false;
            for (int i = 0; i < classSchedules.Count; i++)
            {
                isDuplicate = false;
                ClassSchedule cs = classSchedules[i];
                if (!cs.Room.Equals(room))
                    continue;
                for (int j = 0; j < cs.SlotSchedules.Count; j++)
                {
                    SlotSchedule ss = cs.SlotSchedules[j];
                    if (ss != null)
                    {
                        string[] listSlot = ss.Slot.Split(";");
                        foreach (string slotString in listSlot)
                        {
                            int slot = Int32.Parse(slotString);
                            if (schedule[slot - 1, j] == 0)
                                schedule[slot - 1, j] = 1;
                            else
                            {
                                isDuplicate = true;
                                break;
                            }
                        }
                    }
                    if (isDuplicate)
                        break;
                }
                if (isDuplicate)
                    listDuplicate.Add(i + 1);
            }
            return listDuplicate;
        }*/
        List<String> GetListClass()
        {
            List<string> listClass = new List<string>();
            foreach (ClassSchedule cs in classSchedules)
            {
                if (!listClass.Contains(cs.StudentGroup))
                {
                    listClass.Add(cs.StudentGroup);
                }
            }
            return listClass;
        }
        List<string> GetListTeacher()
        {
            List<string> listTeacher = new List<string>();
            foreach(ClassSchedule cs in classSchedules)
            {
                foreach(SlotSchedule ss in cs.SlotSchedules)
                {
                    if(ss != null)
                    {
                        if (!listTeacher.Contains(ss.Teacher))
                            listTeacher.Add(ss.Teacher);
                    }
                }
            }
            return listTeacher;
        }

        List<int> CheckTeacher(string teacher)
        {
            List<int> listDuplicate = new List<int>();
            for (int i = 0; i < classSchedules.Count; i++)
            {
                ClassSchedule cs1 = classSchedules[i];
                if (!IsTeacherInList(cs1.SlotSchedules, teacher))
                    continue;
                for(int j = i+1; j < classSchedules.Count; j++)
                {
                    ClassSchedule cs2 = classSchedules[j];
                    if (!IsTeacherInList(cs2.SlotSchedules, teacher))
                        continue;
                    if (IsTeacherDuplicate(cs1, cs2))
                        listDuplicate.Add(j + 2);
                }
            }
            return listDuplicate;
        }
        //co the toi uu lai duoc (hoi phuc tap vi phai so sanh ca teacher va date)
        public bool IsTeacherDuplicate(ClassSchedule cs1, ClassSchedule cs2)
        {
            foreach(ScheduleDetail sd1 in cs1.ScheduleDetails)
            {
                int indexDateEqual = BinarySearchDate(cs2.ScheduleDetails, sd1.Date, 0, cs2.ScheduleDetails.Count - 1);
                if (indexDateEqual != -1)
                {
                    ScheduleDetail sd2 = cs2.ScheduleDetails[indexDateEqual];
                    if (sd1.Teacher.Equals(sd2.Teacher) && sd1.Slot == sd2.Slot)
                        return true;
                }
            }
            return false;
        }

        public bool IsTeacherInList(List<SlotSchedule> schedules, string teacher)
        {
            foreach(SlotSchedule ss in schedules)
            {
                if (ss == null)
                    continue;
                if (ss.Teacher.Equals(teacher))
                    return true;
            }
            return false;
        }
        /*List<int> CheckTeacher(string teacher)
        {
            List<int> listDuplicate = new List<int>();
            int[,] schedule = { { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 },
                                { 0, 0, 0, 0, 0, 0 }};
            bool isDuplicate = false;
            for(int i = 0; i < classSchedules.Count; i++)
            {
                isDuplicate = false;
                ClassSchedule cs = classSchedules[i];
                for (int j = 0; j < cs.SlotSchedules.Count; j++)
                {
                    SlotSchedule ss = cs.SlotSchedules[j];
                    if(ss != null)
                    {
                        if(ss.Teacher.Equals(teacher))
                        {
                            string[] listSlot = ss.Slot.Split(";");
                            foreach(string slotString in listSlot)
                            {
                                int slot = Int32.Parse(slotString);
                                if (schedule[slot - 1, j] == 0)
                                    schedule[slot - 1, j] = 1;
                                else
                                {
                                    isDuplicate = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (isDuplicate)
                        break;
                }
                if (isDuplicate)
                    listDuplicate.Add(i + 1);
            }
            return listDuplicate;
        }*/


        void LoadData(string fileName)
        {
            try
            {
                classSchedules = LoadClassSchedule(fileName);
                //classSchedules = LoadClassScheduleXLS(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Load Class Schedule error: " + ex.Message, "Load data error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw ex;
            }
            /*try
            {
                listDay_Off = LoadDay_Off(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Load Day Off error: " + ex.Message, "Load data error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw ex;
            }*/
        }
        List<ClassSchedule> LoadClassScheduleXLS(string fileName)
        {
            List<ClassSchedule> listClassSchedule = new List<ClassSchedule>();
            FileStream fs = new FileStream(fileName, FileMode.Open);

            // Khởi tạo workbook để đọc
            HSSFWorkbook wb = new HSSFWorkbook(fs);

            // Lấy sheet đầu tiên
            ISheet sheet = wb.GetSheetAt(0);

            // đọc sheet này bắt đầu từ row 1 (0 bỏ vì tiêu đề)
            int rowIndex = 1;


            // nếu vẫn chưa gặp end thì vẫn lấy data
            while (sheet.GetRow(rowIndex).GetCell(0).StringCellValue.ToLower() != "end")
            {
                // lấy row hiện tại
                var nowRow = sheet.GetRow(rowIndex);

                ClassSchedule cs = new ClassSchedule();
                var term = nowRow.GetCell(0).StringCellValue;
                cs.Term = term;
                var room = nowRow.GetCell(7).StringCellValue;
                cs.Room = room;
                cs.SlotSchedules = new List<SlotSchedule>();
                string teacher = nowRow.GetCell(8).StringCellValue;
                string slot = "";
                //slot
                //Day 2
                slot = nowRow.GetCell(1).StringCellValue;
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 2;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = nowRow.GetCell(9).StringCellValue.ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 3
                slot = nowRow.GetCell(2).StringCellValue;
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 3;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = nowRow.GetCell(10).StringCellValue.ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 4
                slot = nowRow.GetCell(3).StringCellValue;
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 4;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = nowRow.GetCell(11).StringCellValue.ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 5
                slot = nowRow.GetCell(4).StringCellValue;
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 5;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = nowRow.GetCell(12).StringCellValue.ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 6
                slot = nowRow.GetCell(5).StringCellValue;
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 6;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = nowRow.GetCell(13).StringCellValue.ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 7
                slot = nowRow.GetCell(6).StringCellValue;
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 7;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = nowRow.GetCell(14).StringCellValue.ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //
                cs.SubjectCode = nowRow.GetCell(15).StringCellValue;
                cs.StudentGroup = nowRow.GetCell(16).StringCellValue.ToUpper();
                cs.Type = Int32.Parse(nowRow.GetCell(17).StringCellValue);
                cs.TotalSlot = Int32.Parse(nowRow.GetCell(18).StringCellValue);
                listClassSchedule.Add(cs);
                // tăng index khi lấy xong
                rowIndex++;
            }
            return listClassSchedule;
        }
        List<ClassSchedule> LoadClassSchedule(string fileName)
        {
            List<ClassSchedule> listClassSchedule = new List<ClassSchedule>();
            //string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            string query = "SELECT * FROM [import_data$]";
            OleDbConnection conn = new OleDbConnection(connString);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            OleDbCommand cmd = new OleDbCommand(query, conn);
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ClassSchedule cs = new ClassSchedule();
                cs.Term = reader.GetValue(0).ToString();
                cs.Room = reader.GetValue(7).ToString();
                cs.SlotSchedules = new List<SlotSchedule>();
                string teacher = reader.GetValue(8).ToString().ToLower();
                string slot = "";
                //slot
                //Day 2
                slot = reader.GetValue(1).ToString();
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 2;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = reader.GetValue(9).ToString().ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 3
                slot = reader.GetValue(2).ToString();
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 3;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = reader.GetValue(10).ToString().ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 4
                slot = reader.GetValue(3).ToString();
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 4;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = reader.GetValue(11).ToString().ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 5
                slot = reader.GetValue(4).ToString();
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 5;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = reader.GetValue(12).ToString().ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 6
                slot = reader.GetValue(5).ToString();
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 6;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = reader.GetValue(13).ToString().ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //Day 7
                slot = reader.GetValue(6).ToString();
                if (slot.Length > 0)
                {
                    SlotSchedule ss = new SlotSchedule();
                    ss.Date = 7;
                    ss.Slot = slot;
                    if (teacher.Length > 0)
                        ss.Teacher = teacher;
                    else
                        ss.Teacher = reader.GetValue(14).ToString().ToLower();
                    cs.SlotSchedules.Add(ss);
                }
                else
                {
                    cs.SlotSchedules.Add(null);
                }
                //
                cs.SubjectCode = reader.GetValue(15).ToString();
                cs.StudentGroup = reader.GetValue(16).ToString().ToUpper();
                cs.Type = Int32.Parse( reader.GetValue(17).ToString());
                cs.TotalSlot = Int32.Parse(reader.GetValue(18).ToString());
                cs.StartDate = reader.GetDateTime(19);
                string startDatePart2 = reader.GetValue(30).ToString();
                if(startDatePart2.Length > 0)
                    cs.StartDatePart2 = reader.GetDateTime(30);
                #region day off
                List<DateTime> listDay_Off = new List<DateTime>();
                List<DateTime?> listDataDate = new List<DateTime?>();
                string date = reader.GetValue(20).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(20));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(21).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(21));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(22).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(22));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(23).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(23));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(24).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(24));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(25).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(25));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(26).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(26));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(27).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(27));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(28).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(28));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(29).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(29));
                else
                    listDataDate.Add(null);
                for (int i = 0; i <= 2; i++)
                {
                    DateTime startDate;
                    DateTime endDate;
                    if (listDataDate[i * 2] != null)
                    {
                        startDate = (DateTime)listDataDate[i * 2];
                        // chuoi ngay nghi
                        if (listDataDate[i * 2 + 1] != null)
                        {
                            endDate = (DateTime)listDataDate[i * 2 + 1];
                            while (startDate <= endDate)
                            {
                                listDay_Off.Add(startDate);
                                startDate = startDate.AddDays(1);
                            }
                        }
                        else // 1 ngay nghi
                        {
                            listDay_Off.Add(startDate);
                        }
                    }
                    else
                        continue;
                }
                cs.ListDayOff = listDay_Off;
                #endregion

                string ordinalSubject = reader.GetValue(31).ToString();
                if (ordinalSubject.Length == 0)
                    cs.OrdinalSubject = 0;
                else
                    cs.OrdinalSubject = Int32.Parse(ordinalSubject);
                listClassSchedule.Add(cs);
            }
            reader.Close();
            conn.Close();
            conn.Dispose();
            return listClassSchedule;
        }
        /*List<DateTime> LoadDay_Off(string fileName)
        {
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            string query = "SELECT * FROM [import_data$]";
            OleDbConnection conn = new OleDbConnection(connString);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            OleDbCommand cmd = new OleDbCommand(query, conn);
            OleDbDataReader reader = cmd.ExecuteReader();
            List<DateTime> listDay_Off = new List<DateTime>();
            List<DateTime?> listDataDate = new List<DateTime?>();
            while (reader.Read())
            {
                string date = reader.GetValue(20).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(20));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(21).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(21));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(22).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(22));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(23).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(23));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(24).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(24));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(25).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(25));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(26).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(26));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(27).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(27));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(28).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(28));
                else
                    listDataDate.Add(null);
                date = reader.GetValue(29).ToString();
                if (date.Length != 0)
                    listDataDate.Add(reader.GetDateTime(29));
                else
                    listDataDate.Add(null);

                break;
            }
            reader.Close();
            conn.Close();
            conn.Dispose();
            for (int i = 0; i <= 2; i++)
            {
                DateTime startDate;
                DateTime endDate;
                if (listDataDate[i * 2] != null)
                {
                    startDate = (DateTime)listDataDate[i * 2];
                    // chuoi ngay nghi
                    if(listDataDate[i*2 +1] != null)
                    {
                        endDate = (DateTime)listDataDate[i * 2 + 1];
                        while(startDate <= endDate)
                        {
                            listDay_Off.Add(startDate);
                            startDate = startDate.AddDays(1);
                        }
                    }
                    else // 1 ngay nghi
                    {
                        listDay_Off.Add(startDate);
                    }
                }
                else
                    continue;
            }
            return listDay_Off;
        }*/
        public string RString(string s, int length)
        {
            if (s == null) throw new Exception("RString: s cannot be NULL.");
            if (s.Length < length) throw new Exception("RString: s.Length < length");
            string tmp = s.Substring(s.Length - length, length);
            return tmp;
        }
        private void btnExportSchedule_Click(object sender, EventArgs e)
        {
            //file name Scheduling_Data_dd_mm_yyyy-hh_min.fap
            string dd = "0" + DateTime.Now.Day.ToString();
            dd = RString(dd, 2);

            string mm = "0" + DateTime.Now.Month.ToString();
            mm = RString(mm, 2);

            string yyyy = DateTime.Now.Year.ToString();

            string hh = "0" + DateTime.Now.Hour.ToString();
            hh = RString(hh, 2);

            string min = "0" + DateTime.Now.Minute.ToString();
            min = RString(min, 2);

            string msg = "Export scheduling data finished!\r\nData files:\r\n";
            #region data csv
            string fName = "Scheduling_Data_" + dd + "_" + mm + "_" + yyyy + "-" + hh + "_" + min + ".csv";
            FileStream fs = new FileStream(fName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine("term,day,slot,room,teacher,subjectcode,sessionnumber,studentgroup1,studentgroup2,studentgroup3,Note,SlotType,Type");
            foreach (ClassSchedule cs in this.classSchedules)
            {
                for(int i = 0; i < cs.ScheduleDetails.Count; i++)
                {
                    ScheduleDetail sd = cs.ScheduleDetails[i];
                    string line = cs.Term + "," + sd.Date.ToString("yyyy-MM-dd") + "," + sd.Slot + "," + cs.Room + "," + sd.Teacher + "," + cs.SubjectCode + "," + (i + 1) + ","
                        + cs.StudentGroup + ",,,,," + cs.Type;
                    sw.WriteLine(line);
                }
                /*List<ScheduleDetail> listScheduleDetail = GetListScheduleDetail(cs, cs.StartDate);
                for (int i = 0; i < listScheduleDetail.Count; i++)
                {
                    ScheduleDetail sd = listScheduleDetail[i];
                    string line = cs.Term + "," + sd.Date.ToString("yyyy-MM-dd") + "," + sd.Slot + "," + cs.Room + "," + sd.Teacher + "," + cs.SubjectCode + "," + (i+1) + ","
                        + cs.StudentGroup + ",,,,," +cs.Type;
                    sw.WriteLine(line);
                }*/
            }
            sw.Flush();
            sw.Close();
            fs.Close();
            msg += "1)" + fName + "\r\n";
            #endregion

            #region data xls
            // khởi tạo wb rỗng
            HSSFWorkbook wb1 = new HSSFWorkbook();

            // Tạo ra 1 sheet
            ISheet sheet1 = wb1.CreateSheet();

            // Bắt đầu ghi lên sheet

            // Tạo row
            var row0 = sheet1.CreateRow(0);

            // Ghi tên cột ở row 1
            var row1 = sheet1.CreateRow(0);
            row1.CreateCell(0).SetCellValue("term");
            row1.CreateCell(1).SetCellValue("day");
            row1.CreateCell(2).SetCellValue("slot");
            row1.CreateCell(3).SetCellValue("room");
            row1.CreateCell(4).SetCellValue("teacher");
            row1.CreateCell(5).SetCellValue("subjectcode");
            row1.CreateCell(6).SetCellValue("sessionnumber");
            row1.CreateCell(7).SetCellValue("studentgroup1");
            row1.CreateCell(8).SetCellValue("studentgroup2");
            row1.CreateCell(9).SetCellValue("studentgroup3");
            row1.CreateCell(10).SetCellValue("Note");
            row1.CreateCell(11).SetCellValue("SlotType");
            row1.CreateCell(12).SetCellValue("Type");

            int row = 1;
            foreach (ClassSchedule cs in this.classSchedules)
            {
                List<ScheduleDetail> listScheduleDetail = cs.ScheduleDetails;

                for (int i = 0; i < listScheduleDetail.Count; i++)
                {
                    // tao row mới
                    var newRow = sheet1.CreateRow(row);
                    row++;
                    // set giá trị
                    ScheduleDetail sd = listScheduleDetail[i];
                    newRow.CreateCell(0).SetCellValue(cs.Term);
                    newRow.CreateCell(1).SetCellValue(sd.Date.ToString("yyyy-MM-dd"));
                    newRow.CreateCell(2).SetCellValue(sd.Slot);
                    newRow.CreateCell(3).SetCellValue(cs.Room);
                    newRow.CreateCell(4).SetCellValue(sd.Teacher);
                    newRow.CreateCell(5).SetCellValue(cs.SubjectCode);
                    newRow.CreateCell(6).SetCellValue(i + 1);
                    newRow.CreateCell(7).SetCellValue(cs.StudentGroup);
                    newRow.CreateCell(8).SetCellValue("");
                    newRow.CreateCell(9).SetCellValue("");
                    newRow.CreateCell(10).SetCellValue("");
                    newRow.CreateCell(11).SetCellValue("");
                    newRow.CreateCell(12).SetCellValue(cs.Type);
                }
            }

            // xong hết thì save file lại
            fName = "Scheduling_Data_" + dd + "_" + mm + "_" + yyyy + "-" + hh + "_" + min + ".xls";
            msg += "2)" + fName + "\r\n";
            FileStream fs2 = new FileStream(fName, FileMode.CreateNew);
            wb1.Write(fs2);
            wb1.Close();
            fs2.Close();
            #endregion
            XSSFWorkbook wb2 = new XSSFWorkbook();

            // Tạo ra 1 sheet
            ISheet sheet2 = wb2.CreateSheet();

            // Bắt đầu ghi lên sheet

            // Tạo row
            row0 = sheet2.CreateRow(0);

            // Ghi tên cột ở row 1
            row1 = sheet2.CreateRow(0);
            row1.CreateCell(0).SetCellValue("term");
            row1.CreateCell(1).SetCellValue("day");
            row1.CreateCell(2).SetCellValue("slot");
            row1.CreateCell(3).SetCellValue("room");
            row1.CreateCell(4).SetCellValue("teacher");
            row1.CreateCell(5).SetCellValue("subjectcode");
            row1.CreateCell(6).SetCellValue("sessionnumber");
            row1.CreateCell(7).SetCellValue("studentgroup1");
            row1.CreateCell(8).SetCellValue("studentgroup2");
            row1.CreateCell(9).SetCellValue("studentgroup3");
            row1.CreateCell(10).SetCellValue("Note");
            row1.CreateCell(11).SetCellValue("SlotType");
            row1.CreateCell(12).SetCellValue("Type");

            row = 1;
            foreach (ClassSchedule cs in this.classSchedules)
            {
                List<ScheduleDetail> listScheduleDetail = cs.ScheduleDetails;

                for (int i = 0; i < listScheduleDetail.Count; i++)
                {
                    // tao row mới
                    var newRow = sheet2.CreateRow(row);
                    row++;
                    // set giá trị
                    ScheduleDetail sd = listScheduleDetail[i];
                    newRow.CreateCell(0).SetCellValue(cs.Term);
                    newRow.CreateCell(1).SetCellValue(sd.Date.ToString("yyyy-MM-dd"));
                    newRow.CreateCell(2).SetCellValue(sd.Slot);
                    newRow.CreateCell(3).SetCellValue(cs.Room);
                    newRow.CreateCell(4).SetCellValue(sd.Teacher);
                    newRow.CreateCell(5).SetCellValue(cs.SubjectCode);
                    newRow.CreateCell(6).SetCellValue(i + 1);
                    newRow.CreateCell(7).SetCellValue(cs.StudentGroup);
                    newRow.CreateCell(8).SetCellValue("");
                    newRow.CreateCell(9).SetCellValue("");
                    newRow.CreateCell(10).SetCellValue("");
                    newRow.CreateCell(11).SetCellValue("");
                    newRow.CreateCell(12).SetCellValue(cs.Type);
                }
            }

            // xong hết thì save file lại
            fName = "Scheduling_Data_" + dd + "_" + mm + "_" + yyyy + "-" + hh + "_" + min + ".xlsx";
            msg += "3)" + fName + "\r\n";
            FileStream fs3 = new FileStream(fName, FileMode.CreateNew);
            wb2.Write(fs3);
            wb2.Close();
            fs3.Close();
            #region tao file xlsx

            #endregion
            MessageBox.Show(msg, "Export data");
        }

        public List<ScheduleDetail> GetListScheduleDetail(ClassSchedule cs, DateTime startDate)
        {
            List<ScheduleDetail> listScheduleDetail = new List<ScheduleDetail>();
            int dayMore = startDate.DayOfWeek - DayOfWeek.Monday;
            DateTime startFirstWeek = startDate.AddDays(-dayMore);
            int[] dateSchedule = {1, 1, 1, 1, 1, 1};
            for(int i = 0; i < 6; i++)
            {
                if (cs.SlotSchedules[i] == null)
                    dateSchedule[i] = 0;
            }
            int numberSlot = GetSlot(cs.SlotSchedules).Split(";").Length;

            int slotCount = 0;
            int startSlotChange = 0;
            int endSlot;
            if (cs.TotalSlot % numberSlot != 0)
            {
                if (cs.TotalSlot == 40 || cs.TotalSlot == 80)
                {
                    startSlotChange = 13;
                    endSlot = 14;
                }
                else
                {
                    startSlotChange = 7;
                    endSlot = 7;
                }
            }
            else
            {
                if(cs.TotalSlot == 80)
                    endSlot = cs.TotalSlot / (2 * numberSlot);
                else
                    endSlot = cs.TotalSlot / numberSlot;
            }

            for(int i = 0; ;i++)
            {
                DateTime startWeek = startFirstWeek.AddDays(i * 7);
                int[] arrayTest = { 1, 1, 1, 1, 1, 1 };
                for(int j = 0; j < 6; j++)
                {
                    if (startWeek.AddDays(j) >= startDate)
                        if (dateSchedule[j] == 1)
                            if (!cs.ListDayOff.Contains(startWeek.AddDays(j)))
                                arrayTest[j] = 1;
                            else
                                arrayTest[j] = 0;
                        else
                            arrayTest[j] = 0;
                    else
                        arrayTest[j] = 0;
                }

                for(int j = 0; j < 6; j++)
                {
                    if(arrayTest[j] == 1)
                    {
                        slotCount++;
                        
                        if(startSlotChange != 0 && slotCount >= startSlotChange)
                        {
                            DateTime date = startWeek.AddDays(j);
                            string[] arraySlot = GetSlotByDate(cs.SlotSchedules, (int)date.DayOfWeek).Split(";");
                            List<string> slotChange = new List<string>();
                            if (cs.TotalSlot == 40 || cs.TotalSlot == 80)
                            {
                                if (Int32.Parse(arraySlot[0]) >= 4)
                                    slotChange = new List<string> { "4", "5" };
                                else
                                    slotChange = new List<string> { "2", "3" };
                            }
                            else
                            {
                                if (Int32.Parse(arraySlot[0]) >= 4)
                                    slotChange = new List<string> { "4", "5", "6" };
                                else
                                    slotChange = new List<string> { "1", "2", "3" };
                            }
                            foreach (string slot in slotChange)
                            {
                                ScheduleDetail sd = new ScheduleDetail();
                                sd.Date = date;
                                sd.Teacher = cs.SlotSchedules[j].Teacher;
                                sd.Slot = slot;
                                listScheduleDetail.Add(sd);
                            }
                        }
                        else
                        {
                            DateTime date = startWeek.AddDays(j);
                            int dateNumber = (int)date.DayOfWeek;
                            string[] arraySlot = GetSlotByDate(cs.SlotSchedules, dateNumber).Split(";");
                            foreach (string slot in arraySlot)
                            {
                                ScheduleDetail sd = new ScheduleDetail();
                                sd.Date = startWeek.AddDays(j);
                                sd.Teacher = cs.SlotSchedules[j].Teacher;
                                sd.Slot = slot;
                                listScheduleDetail.Add(sd);
                            }
                        }
                    }
                    if (slotCount == endSlot)
                        break;
                }
                if (slotCount == endSlot)
                    break;
            }
            return listScheduleDetail;
        }
        public string GetSlotByDate (List<SlotSchedule> listSS, int day)
        {
            SlotSchedule ss =  listSS[day - 1];
            if (ss != null)
                return ss.Slot;
            return null;
        }
        public string GetSlot(List<SlotSchedule> listSS)
        {
            foreach (SlotSchedule ss in listSS)
            {
                if (ss != null)
                    return ss.Slot;
            }
            return null;
        }

    }
}
