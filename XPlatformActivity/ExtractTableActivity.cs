using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Activities;
using System.ComponentModel;
using UiPath.Core;
using System.Drawing;

namespace XPlatformActivity
{

    class GridHeaderName
    {
        readonly public int X;
        readonly public int Width;
        readonly public string name;

        public GridHeaderName(int x, int width, string name)
        {
            this.X = x;
            this.Width = width;
            this.name = name;
        }
    }

    public sealed class ExtractTable : CodeActivity
    {
        // 선택된 Element로 Find Element Activity로 선택된 항목을 의미함 
        [Category("Input")]
        [RequiredArgument]
        public InArgument<UiElement> Element { get; set; }

        // 사용자가 입력한 table head 값 
        [Category("Input")]
        public InArgument<string[]> UserHeaderNames { get; set; }

        // 중첩된 헤더가 있는 경우 제거할 이름 리스트 
        [Category("Input")]
        public InArgument<string[]> ExcludeHeaderNames { get; set; }

        // 추출된 결과 테이블을 의미 
        [Category("Output")]
        [RequiredArgument]
        public OutArgument<DataTable> Table { get; set; }

        // Execute 메서드에서 값을 반환합니다.
        private DataTable table;

        // table의 role type이 row 이거나 outline 이거나 
        private string []roles = { "row", "outline item" };
        private string NO_NAME_COLULMN = "Column{0}";

        private string GetSelectorType( string hdr, string role)
        {
            return String.Format("<ctrl name='{0}' role='{1}' />", hdr, role);
        }

        private bool IsInside(List<GridHeaderName> rects, System.Drawing.Rectangle cur, out string name)
        {
            bool inside = false;
            name = String.Empty;
            foreach(GridHeaderName r in rects)
            {
                if ( (cur.X >= r.X && cur.Width + cur.X < r.Width + r.X) || //시작할때는 cur.X == r.X 
                    (cur.X > r.X && cur.Width + cur.X <= r.Width + r.X))  // 끝날때는 cur.Width + cur.X == r.With + r.X  
                {
                    inside = true;
                    name = r.name;
                    break;
                }  
            }
            return inside;
        }

        private bool BuildTableHeaderFromUserInput( DataTable table, string[] colNames)
        {
            bool result = true;

            foreach( string name in colNames)
            {
                DataColumn column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = name;
                table.Columns.Add(column);
            }

            return result;
        }

        //UiElement를 탐색해서 테이블 헤더를 구성한다. 
        private bool BuildTableHeaderFromUiElement( DataTable table,  UiElement [] rows, List<string> unusedHdrNames)
        {
            bool result = true;
            List<GridHeaderName> rects = new List<GridHeaderName>();
            List<string> hdrNames = new List<string>();
            foreach ( UiElement row in rows)
            {
                UiElement[] hdrs = row.FindAll(FindScope.FIND_DESCENDANTS, new Selector("<ctrl role='text' />"));
                int columnIndex = 0;

                //중첩이 있는 경우 포함하고 있는 header name을 꺼내고 이 이름을 나중에 사용하지 않도록 한다. 
                foreach (UiElement hdr in hdrs)
                {
                    string excHdrName = String.Empty;
                    Rectangle r = hdr.GetAbsolutePosition();
                    bool donotuse = false;
                    if (rects.Count > 0)
                    {
                        //중첩이 있는지 체크 
                        if ((donotuse = IsInside(rects, r, out excHdrName)))
                        {
                            //여기에 포함된 이름은 사용하지 않아 한다. 
                            if(!unusedHdrNames.Contains<string>(excHdrName))
                            {
                                unusedHdrNames.Add(excHdrName);
                                //System.Console.Out.WriteLine("이 컬럼:{0} 중첩으로 판단 ", excHdrName);
                            }
                        }
                    }
                    string name = hdr.Get("name").ToString().Trim();
                    rects.Add(new GridHeaderName(r.X, r.Width, name));
                    hdrNames.Add(name);
                }
                foreach( string name in hdrNames)
                {
                    DataColumn column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = name;
                    if( ! unusedHdrNames.Contains<string>( name))
                        table.Columns.Add(column);
                }
            }

            return result;
        }

        //UiElement 에서 찾은 내용으로 Table 본문을 만든다. 이때 앞에서 찾은 중첩된 이름이 있다면 이 이름은 제거하고 사용한다. 
        private bool BuildTableBodyFromUiElement( DataTable table, UiElement []rows ,List<string> unusedHdrNames)
        {
            bool result = true;

            foreach (UiElement row in rows) // body 1 개 
            {
                UiElement[] bodies = row.FindAll(FindScope.FIND_CHILDREN, new Selector("<ctrl role='text'/>"));
                int bodyIndex = 0;
                DataRow dataRow = null;
                foreach (UiElement body in bodies)
                {
                    bool nameFixed = false;
                    int colIndex = bodyIndex % table.Columns.Count;
                    if (colIndex == 0) // 처음이면 
                    {
                        dataRow = table.NewRow();
                    }
                    string name = body.Get("name").ToString().Trim();
                    System.Console.Out.WriteLine("현재 셀 값 : {0}", name);
                    foreach (string excName in unusedHdrNames)
                    {
                        int exIdx = name.IndexOf(excName);
                        if (exIdx > 0)
                        {
                            dataRow[colIndex] = name.Substring(0, exIdx - 1).Trim();
                            nameFixed = true;
                            //System.Console.Out.WriteLine("current row - name: {0}, dataRow: {1}", name, dataRow[colIndex]);
                        } 
                        else if( exIdx == 0) // 해당 셀에 값이 없다면 
                        {
                            dataRow[colIndex] = "";
                            nameFixed = true;
                        }
                    }
                    if (!nameFixed) // 이름이 확정되지 않았다면 
                    {
                        int idx = name.IndexOf(table.Columns[colIndex].ColumnName);
                        if (idx > 0) // 뒤에 컬럼 이름이 있다는 것이고 
                        {
                            dataRow[colIndex] = name.Substring(0, idx).Trim();
                        }
                        else if (table.Columns[colIndex].ColumnName.StartsWith("Column")) // 이름이 없었다면 전부를 다 사용하고 
                        {
                            dataRow[colIndex] = name;
                        }
                    }
                    //마지막 이면 만들어진 DataRow를 Table에 추가 
                    if (colIndex == this.table.Columns.Count - 1) 
                    {
                        table.Rows.Add(dataRow);
                    }
                    bodyIndex++;
                }
            }
            return result;
        }


        protected override void Execute(CodeActivityContext context)
        {
            this.table = new DataTable();
            bool roleMatched = false;
            List<string> unusedHdrNames = new List<string>();

            string[] userHeaderNames = UserHeaderNames.Get(context);

            UiElement _table = Element.Get<UiElement>(context);
            if ( _table == null)
            {
                throw new ArgumentException("Element can not be null, please check input Element value");
            }

            foreach (string role in roles)
            {
                if (roleMatched) //이미 데이터를 추출했다면 다음 role에 대해서 검사하지 않도록 
                    break;
                UiElement[] rows = _table.FindAll(FindScope.FIND_CHILDREN, new UiPath.Core.Selector( GetSelectorType( "head", role)));
                if (rows == null)
                {
                    System.Console.Out.WriteLine(" Find Children with <ctrl name='{0}' role='{1}' /> was failed, no child element exist", "head", role);
                    continue;
                }
                roleMatched = true;
                //우선 테이블 헤더부분을 만들어 내고 
                if( userHeaderNames is null )
                {
                    BuildTableHeaderFromUiElement(this.table, rows, unusedHdrNames);
                }
                else
                {
                    BuildTableHeaderFromUserInput(this.table, userHeaderNames);
                    string[] excNames = ExcludeHeaderNames.Get(context);
                    if( ! (excNames is null))
                        foreach( string name in excNames)
                        {
                            unusedHdrNames.Add(name);
                        }
                }
                //테이블 본문을 만들어 낸나. 
                rows = _table.FindAll(FindScope.FIND_CHILDREN, new UiPath.Core.Selector(GetSelectorType("body", role)));
                BuildTableBodyFromUiElement(this.table, rows, unusedHdrNames);
            }
            // 최종 결과 table 할당
            Table.Set(context, this.table);
        }
    }
}
