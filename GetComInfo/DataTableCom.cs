using System.Data;

namespace GetComInfo
{
    internal class DataTableCom
    {
        public DataTable dataTableCom() => new DataTable()
        {
            Columns =
            {
                {
                  "COM",
                  typeof (string)
                },
                {
                  "Check",
                  typeof (bool)
                },
                {
                  "Status",
                  typeof (string)
                },
                {
                  "Numberphone",
                  typeof (string)
                },
                {
                  "Content",
                  typeof (string)
                },
                {
                  "Description",
                  typeof (string)
                }
            }
        };
    }
}
