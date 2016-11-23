using NPOI.HSSF.UserModel;
using NPOI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data.SqlClient;

namespace ListData
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("开始导入数据。");
            //1.读取excel数据
            ReadNPOI();
            //2.处理业务逻辑，封闭Model

            //3.循环将每个Model写入数据库
            Console.WriteLine("导入完成。");
            Console.ReadKey();

            //说明：因数据量只有150条，故可一次性批量处理业务逻辑 及 循环写入数据库
        }
        //读取xls文件
        public static void ReadNPOI()
        {
            StringBuilder sbr = new StringBuilder();
            List<ListModel> list = new List<ListModel>();
            using (FileStream fs = File.OpenRead(@"powerlist.xlsx"))   //打开myxls.xls文件
            {
                XSSFWorkbook wk = new XSSFWorkbook(fs);   //把xls文件中的数据写入wk中
                for (int i = 0; i < wk.NumberOfSheets; i++)  //NumberOfSheets是myxls.xls中总共的表数
                {
                    ISheet sheet = wk.GetSheetAt(i);   //读取当前表数据

                    for (int j = 0; j <= sheet.LastRowNum; j++)  //LastRowNum 是当前表的总行数
                    {
                        IRow row = sheet.GetRow(j);  //读取当前行数据
                        if (row != null)
                        {
                            //sbr.Append("-------------------------------------\r\n"); //读取行与行之间的提示界限
                            ListModel model = new ListModel();
                            for (int k = 0; k <= row.LastCellNum; k++)  //LastCellNum 是当前行的总 列 数                            
                            {
                                ICell cell = row.GetCell(k);  //当前表格                                
                                //if (cell != null)
                                //{
                                //    //sbr.Append(cell.ToString());   //获取表格中的数据并转换为字符串类型
                                //    //
                                //}

                                switch (k)
                                {
                                    case 0://地区，如郑州
                                        //通过城市名查询得到：城市ID      
                                        int id = -1;
                                        object o = GetCityIdByName(cell.ToString());
                                        if (o != null)
                                        {
                                            int.TryParse(GetCityIdByName(cell.ToString()).ToString(), out id);
                                        }                                        
                                        model.CityId = id;
                                        //将得到的城市ID写入数据库
                                        break;
                                    case 1://公司名
                                        model.CompanyName = cell.ToString();
                                        break;
                                    case 2://主营类别
                                        //处理方式为：将所有多个主营的企业删除，随后手动添加即可。
                                        model.TypeId = int.Parse(GetTypeIdByName(cell.ToString()).ToString());
                                        break;
                                    case 3://主营详细分类
                                        model.CompanyMain = cell.ToString();
                                        break;
                                    case 4://钢厂
                                        //此cell不添加进数据库
                                        break;
                                    case 5://联系人姓名
                                        //不加入数据库
                                        break;
                                    case 6://手机号
                                        //读取后写入
                                        model.CompanyTel = cell.ToString();
                                        break;
                                }
                            }
                            if(!string.IsNullOrEmpty(model.CompanyMain)&&!string.IsNullOrEmpty(model.CompanyName)&&!string.IsNullOrEmpty(model.CompanyTel))
                            {
                                list.Add(model);
                                Console.WriteLine("加入实体中...");

                            }

                        }
                    }
                }
            }
            //sbr.ToString();
            //using (StreamWriter wr = new StreamWriter(new FileStream(@"c:/myText.txt", FileMode.Append)))  //把读取xls文件的数据写入myText.txt文件中
            //{
            //    wr.Write(sbr.ToString());
            //    wr.Flush();
            //}

            //遍历list，将其插入数据库
            foreach (var m in list)
            {
                AddModel(m);
                Console.WriteLine("插入数据中...");
            }
        }

        protected static object GetCityIdByName(string name)
        {
            string strSql = "select Id from [dbo].[ListCity] where CityName like '%" + name + "%'";
            //SqlParameter pm = new SqlParameter("@Name", name);            
            return SqlHelper.ExecuteScalar(strSql);
        }
        protected static object GetTypeIdByName(string name)
        {
            string strSql = "select Id from [dbo].[ListProductType] where ProductType like '%" + name + "%'";
            //SqlParameter pm = new SqlParameter("@Name", name);
            //            object o = SqlHelper.ExecuteScalar(strSql, pm);
            return SqlHelper.ExecuteScalar(strSql);
        }
        protected static void AddModel(ListModel model)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("insert into ListCompany(CityId,TypeId,CompanyName,CompanyTel,CompanyMain) values(");
            sb.Append(model.CityId+",");
            sb.Append(model.TypeId + ",'");
            sb.Append(model.CompanyName + "','");
            sb.Append(model.CompanyTel + "','");
            sb.Append(model.CompanyMain + "')");
            SqlHelper.ExecuteNonQuery(sb.ToString());
        }

    }
    public class ListModel
    {
        public int CityId { get; set; }
        public int TypeId { get; set; }
        public string CompanyName { get; set; }
        public string CompanyTel { get; set; }
        public string CompanyMain { get; set; }
    }
    class SqlHelper
    {
        
        private static readonly string STRCONN = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
        public static SqlDataReader GetReader(string strSql, params SqlParameter[] pms)
        {
            SqlConnection conn = new SqlConnection(STRCONN);

            using (SqlCommand cmd = new SqlCommand(strSql, conn))
            {
                if (pms != null)
                {
                    cmd.Parameters.AddRange(pms);
                }
                conn.Open();
                return cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
            }

        }

        public static int ExecuteNonQuery(string strSql, params SqlParameter[] pms)
        {
            using (SqlConnection conn = new SqlConnection(STRCONN))
            {
                using (SqlCommand cmd = new SqlCommand(strSql, conn))
                {
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }
                    conn.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }
        public static object ExecuteScalar(string strSql, params SqlParameter[] pms)
        {
            using (SqlConnection conn = new SqlConnection(STRCONN))
            {
                using (SqlCommand cmd = new SqlCommand(strSql, conn))
                {
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }
                    conn.Open();
                    return cmd.ExecuteScalar();
                }
            }
        }

    }
}
