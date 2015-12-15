using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data.SQLite;
using System.Data;
namespace GetBusAndBikeOfSz
{
    class Program
    {
        static ArrayList GetLinesFromExcel(string filename)
        {
            ArrayList lines = new ArrayList();
            Application  excel = new Application();
            Workbook wb =  excel.Workbooks.Open(filename);
            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
                Worksheet ws = wb.Worksheets.get_Item(i);
                int rowcnt = ws.UsedRange.Cells.Rows.Count;
                int colcnt = ws.UsedRange.Cells.Columns.Count;

                //get line number col
                int line_number_col = 0;
                //string cell_content = "";
                for (int j = 1; j <= colcnt; j++)
                {
                    Range rng = ws.Cells[1, j];
                    if (string.Compare(rng.Text, "【番号】") == 0)
                    {
                        line_number_col = j;
                        break;
                    }
                }
                if (line_number_col <= 0)
                    continue;
                for (int j = 2; j <= rowcnt; j++)
                {
                    Range rng = ws.Cells[j, line_number_col];
                    if (!string.IsNullOrEmpty(rng.Text))
                        lines.Add(rng.Text);
                }
            }

            excel.Workbooks.Close();
            excel.Quit();
            return lines;
        }
        public class Line
        {
            public string lcompanyGuid;
            public string lesname;
            public string lguid;
            public string lfstdftime;
            public string lname;
            public string ldirection;
            public string lfsname;
            public string lfstdetime;
            public string regionId;
            public string ldistance;
        }
        static ArrayList GetMoreLineInfo(ArrayList linesname)
        {
            ArrayList lines = new ArrayList();
            for(int i=0;i<linesname.Count;i++){
                string url = string.Format("http://wap.139sz.cn/cx/pp/searchLinesByName.php?lname={0}",linesname[i].ToString());
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Method = "GET";
                request.Accept = "application/json";
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; rv:42.0) Gecko/20100101 Firefox/42.0";
                request.Referer = "http://wap.139sz.cn/cx/";
                request.Host = "wap.139sz.cn";
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;

                //如果主体信息不为空，则接收主体信息内容
                //if (response.ContentLength <= 0)
                //   return;
                //接收响应主体信息
                using (Stream stream = response.GetResponseStream())
                {
                    StreamReader streamRead = new StreamReader(stream);
                    string responseString = streamRead.ReadToEnd();

                    JObject jo = (JObject)JsonConvert.DeserializeObject(responseString);
                    bool bres = Convert.ToBoolean(jo["success"].ToString());
                    if (!bres)
                    {
                        Console.WriteLine(linesname[i].ToString());
                    }
                    else
                    {
                        JArray ja = (JArray)jo["busInfo"];
                        for (int j = 0; j < ja.Count; j++)
                        {
                            string lguid = ja[j]["lguid"].ToString();
                            bool bAdd = true;
                            for (int k = 0; k < lines.Count; k++)
                            {
                                Line l = (Line)lines[k];
                                if (string.Compare(l.lguid, lguid) == 0)
                                {
                                    bAdd = false;
                                    break;
                                }
                            }
                            if (!bAdd)
                                continue;
                            Line ln = JsonConvert.DeserializeObject<Line>(ja[j].ToString());
                            lines.Add(ln);
                        }
                    }
                }
            }
            return lines;
        }
        static void WriteLineToDb(string dbname,ArrayList buses)
        {
            string dbpath = "Data Source =" + dbname;
            using (SQLiteConnection conn = new SQLiteConnection(dbpath))
            {
                try
                {
                    conn.Open();
                    string sql = @"Create table if not exists lines(id int not null,
                                                                name varchar(100) not null,
                                                                direction varchar(100) not null,
                                                                guid varchar(20) not null,
                                                                info blob,
                                                                PRIMARY KEY(id))";
                    SQLiteCommand cmdCreateTable = new SQLiteCommand(sql, conn);
                    cmdCreateTable.ExecuteNonQuery();//如果表不存在，创建数据表 

                    using (SQLiteTransaction tran = conn.BeginTransaction())
                    {
                        for (int i = 0; i < buses.Count; i++)
                        {
                            Line l = (Line)buses[i];
                            SQLiteCommand cmd = new SQLiteCommand(conn);
                            cmd.Transaction = tran;
                            cmd.CommandText = "insert into lines(id,name,direction,guid,info) values(@lid,@lname,@ldirection,@lguid,@linfo)";
                            string info = "{\"time\":\"" + l.lfstdftime + "<-->" + l.lfstdetime + "\"}";
                            byte[] byteArray = System.Text.Encoding.ASCII.GetBytes ( info );
                            cmd.Parameters.AddRange(new[]{
                                new SQLiteParameter("lid",i),
                                new SQLiteParameter("lname",l.lname),
                                new SQLiteParameter("ldirection",l.ldirection),
                                new SQLiteParameter("lguid",l.lguid)/*,
                                new SQLiteParameter("linfo",DbType.Binary).Value = byteArray*/
                            });
                            SQLiteParameter param = new SQLiteParameter();
                            param.DbType = DbType.Binary;
                            param.ParameterName = "linfo";
                            param.Value = byteArray;
                            cmd.Parameters.Add(param);
                            cmd.ExecuteNonQuery();
                        }
                        tran.Commit();
                    }

                    conn.Close();
                }
                catch (Exception e)
                {
                    string txt = e.Source + e.Message;
                    Console.WriteLine(txt);
                }
            }
        }
        static void test()
        {
            for (int i = 0; i < 1000000; i++)
            {
                try
                {
                    string url = "http://content.2500city.com/api18/bus/getLineInfo?sign=c0d9d732020ac85fad19fefc849475f4&uid=1443143&lng=120.746672&Guid=bb6745f8-8b8d-481c-9989-a0bbd3a3cb42&lat=31.316076";
                    HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                    request.Method = "GET";
                    //request.Accept = "application/json";
                    //request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; rv:42.0) Gecko/20100101 Firefox/42.0";
                    //request.Referer = "http://wap.139sz.cn/cx/";
                    //request.Host = "wap.139sz.cn";
                    HttpWebResponse response = request.GetResponse() as HttpWebResponse;

                    //如果主体信息不为空，则接收主体信息内容
                    //if (response.ContentLength <= 0)
                    //   return;
                    //接收响应主体信息
                    using (Stream stream = response.GetResponseStream())
                    {
                        StreamReader streamRead = new StreamReader(stream);
                        string responseString = streamRead.ReadToEnd();
                        if (string.IsNullOrEmpty(responseString))
                            Console.WriteLine("Empty :{0}", i);
                        JObject jo = (JObject)JsonConvert.DeserializeObject(responseString);
                        int bres = Convert.ToInt32(jo["errorCode"].ToString());
                        if (bres != 0)
                        {
                            string msg = jo["errorMsg"].ToString();
                            Console.WriteLine("Error :{0}, msg:{2}", i, msg);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        static void usage()
        {
            string prompt = @"you can input following commands:\n
                              \t-x:the input excel,it save all lines of city;\n
                              \t-d:the output db";
            Console.Write(prompt);
        }
        static void Main(string[] args)
        {
            //test();
            if (args.Length == 0)
            {
                usage();
                return;
            }
            string xlsname = "", dbname = "";
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-x":
                        xlsname = args[i + 1].ToString(); i++;
                        break;
                    case "-d":
                        dbname = args[i + 1].ToString(); i++;
                        break;
                    case "-h":
                        usage();
                        break;
                    default: break;
                }
            }
            ArrayList linesname = GetLinesFromExcel(xlsname);
            ArrayList lines /*= new ArrayList();//*/= GetMoreLineInfo(linesname);
            Console.WriteLine("lines count: {0}", lines.Count);
            WriteLineToDb(dbname, lines);
            Console.WriteLine("Finish");
            Console.ReadLine();
        }
    }
}
