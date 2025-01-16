namespace ExcelTester.Attributes
{
    /// <summary>
    /// 做一個 Excel 的欄位對應
    /// </summary>
    public class ExcelColumnMapping : Attribute
    {
        public string Column { get; set; }

        public ExcelColumnMapping(string column)
        {
            this.Column = column.ToUpper();
        }

        public int ColumnIndex
        {
            get
            {
                double num = 0;
                for (int i = 0; i < this.Column.Length; i++)
                    num += (this.Column[i] - 64) * Math.Pow(26, this.Column.Length - 1 - i);

                return Convert.ToInt32(num);
            }
        }
    }
}

