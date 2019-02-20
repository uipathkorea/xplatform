using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Activities;
using System.ComponentModel;
using UiPath.Core;

namespace XPlatformActivity
{

    public sealed class ExtractTable : CodeActivity
    {
        // 선택된 Element로 Find Element Activity로 선택된 항목을 의미함 
        [Category("Input")]
        public InArgument<UiElement> Element { get; set; }


        // 추출된 결과 테이블을 의미 
        [Category("Output")]
        public OutArgument<DataTable> Table { get; set; }

        // Execute 메서드에서 값을 반환합니다.
        private DataTable table;

        protected override void Execute(CodeActivityContext context)
        {
            this.table = new DataTable();

            UiElement _table = Element.Get<UiElement>(context);
            System.Console.Out.WriteLine("Get Element - Table element ");
            if ( _table == null)
            {
                System.Console.Out.WriteLine("Parameter Element is null");
                return;
            }
            System.Console.Out.WriteLine("Check to table's row children...");
            UiElement []rows = _table.FindAll(FindScope.FIND_CHILDREN, new UiPath.Core.Selector("<ctrl name='head' role='row' />"));
            if( rows == null)
            {
                System.Console.Out.WriteLine("_table.FindAll is null");
                return;
            }
            System.Console.Out.WriteLine(" head.row count is {0}", rows.Length);
            int rowCount = rows.Length;

            foreach (UiElement row in rows) // first row is header 
            {

                UiElement[] hdrs = row.FindAll(FindScope.FIND_DESCENDANTS, new Selector("<ctrl role='text' />"));
                if (hdrs == null)
                {
                    System.Console.Out.WriteLine(" row.FindAll is null");
                    return;
                }
                System.Console.Out.WriteLine("header count {0}", hdrs.Length);
                foreach (UiElement hdr in hdrs)
                {
                    DataColumn column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = hdr.Get("name").ToString().Trim().Split().Last<string>();
                    if( ! "전체".Equals(column.ColumnName))
                    {
                        this.table.Columns.Add(column);
                    }
                }
            }
            System.Console.Out.WriteLine(" Header 추가 완료 ");

            rows = _table.FindAll(FindScope.FIND_CHILDREN, new UiPath.Core.Selector("<ctrl name='body' role='row' />"));
            foreach (UiElement row in rows) // body 1 개 
            { 
                UiElement []bodies = row.FindAll(FindScope.FIND_CHILDREN, new Selector("<ctrl role='text'/>"));
                if( bodies == null)
                {
                    System.Console.Out.WriteLine("body row.FindAll is null");
                    return;
                }
                int bodyIndex = 0;
                DataRow dataRow = null;
                foreach( UiElement body in bodies)
                {
                    if ( bodyIndex % this.table.Columns.Count == 0 ) // 처음이면 
                    {
                        dataRow = this.table.NewRow();
                    }
                    String[] items = body.Get("name").ToString().Trim().Split();
                    dataRow[items.Last<string>()] = items.First<string>();

                    if (bodyIndex % this.table.Columns.Count == this.table.Columns.Count - 1) // 마지막 이면 
                    {
                        this.table.Rows.Add(dataRow);
                    }
                    bodyIndex++;   
                }
            }
        
            // 최종 결과 table 할당
            Table.Set(context, this.table);
        }
    }
}
