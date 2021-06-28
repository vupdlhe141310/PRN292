# PRN292




public static bool checkSort(DataTable tbl1, DataTable tbl2)
        {
            if (tbl1.Rows.Count != tbl2.Rows.Count || tbl1.Columns.Count != tbl2.Columns.Count)
                return false;


            for (int i = 0; i < tbl1.Rows.Count; i++)
            {
                for (int c = 0; c < tbl1.Columns.Count; c++)
                {
                    if (!Equals(tbl1.Rows[i][c], tbl2.Rows[i][c]))
                        return false;
                }
            }
            return true;
        }
        //check PE result
        public static bool checkData(DataTable tbl1, DataTable tbl2)
        {
            return true;
        }
        public static Student getPEResult(DataTable[] dtStudentAnswer, DataTable[] dtSolution, string sId, string sName, string sCode)
        {

            string mess = "";
            double totalMark = 0;

            for (int i = 0; i < 10; i++)
            {
                mess += "Q" + (i + 1);
                if (dtStudentAnswer[i] == null)
                {
                    mess += " Empty";
                }
                else
                {
                    double QMark = 0;
                    if (checkData(dtStudentAnswer[i], dtSolution[i]))
                    {
                        mess += "check Data: Passed => ";
                        QMark += 0.5;
                        Console.WriteLine(mess + "+ " + QMark);
                        if (checkSort(dtStudentAnswer[i], dtSolution[i]))
                        {
                            mess += "check Sort: Passed => ";
                            QMark += 0.5;
                            Console.WriteLine(mess + "+ " + QMark);
                        }
                        else
                        {
                            mess += " check Sort: Not Pass";
                        }                      
                    }
                    else
                    {
                        mess += " check Data: Not Pass";
                    }
                    totalMark += QMark;
                    Console.WriteLine("Total Point: " + totalMark);
                }
            }
            return new Student();
        }
