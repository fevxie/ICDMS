using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Excel2Xml
{
    public class BuildXml
    {
        public List<Row> Rows { get; private set; }

        public BuildXml(List<Row> rows)
        {
            this.Rows = rows;
        }

        List<Command> commandList = new List<Command>();

        public void SetXml()
        {
            for (int i = 0; i < Rows.Count; i++)
            {
                if (Convert.ToInt32(Rows[i].RowID) <3)
                {
                    continue;
                }

                commandList.Add(new Command
                {
                    ID = Convert.ToInt32(Rows[i].RowID),
                    Cells = Rows[i].Cells
                });
            }

            XElement doc = new XElement("Commands");

            doc.Save("result1.xml");

            StringBuilder sb = new StringBuilder();

            Dictionary<string, string> collapseList = new Dictionary<string,string>();

            for (int i = 0; i < 12; i++)
            {
                if (commandList[i].Cells.Count() > 7 && sb.Length >0)
                {
                    collapseList.Add(sb.ToString().Split(';').FirstOrDefault(), sb.ToString());
                    sb.Remove(0, sb.Length);
                    sb.Append(commandList[i].Cells[0].Data + ";");
                }
                else
                {
                    sb.Append(commandList[i].Cells[0].Data + ";");
                }

                doc.Add(new XElement("Command",
                     new XAttribute("id", commandList[i].Cells[0].Data),
                     new XElement("Address", commandList[i].Cells[1].Data),
                     new XElement("Inputs",
                         new XElement("Input",
                             new XAttribute("name",commandList[i].Cells[2].Data),
                             new XAttribute("Code",commandList[i].Cells[3].Data),
                             new XAttribute("Type",commandList[i].Cells[4].Data))),
                     new XElement("Condition",
                         new XAttribute("statement", commandList[i].Cells[5].Data),
                         new XAttribute("Value", commandList[i].Cells[6].Data))));

                doc.Save("result1.xml");
            }

            //最后一次循环保存的值
            if (sb.Length>0)
            {
                collapseList.Add(sb.ToString().Split(';').FirstOrDefault(), sb.ToString());
                sb.Remove(0, sb.Length);
            }

            //生成合并项
            foreach (var command in commandList.Where(o => o.Cells.Count() > 7))
            {
                if (command.Cells[0].Data != "EOF")
                {
                    doc.LastNode.AddAfterSelf(new XElement("Collpase",
                          new XElement("CommandID", collapseList[command.Cells[0].Data]),
                          new XElement("Action", command.Cells[7].Data),
                          new XElement("Address", command.Cells[8].Data),
                          new XElement("OutPut",
                              new XAttribute("Name", command.Cells[9].Data),
                              new XAttribute("Code", command.Cells[10].Data),
                              new XAttribute("Type", command.Cells[11].Data),
                              new XAttribute("Value", command.Cells[12].Data)
                              )
                       ));
                    doc.Save("result1.xml");
                }
            }
        }
    }

    public class Command{
        public int ID { get; set; }

        public Cell[] Cells { get; set; }
    }

    public class Input{
        public string Name { get; set; }
        public string Code { get; set; }
        public string Type { get; set; }
        public string Statement { get; set; }
    }

    public class Output{
        public string Name { get; set; }

        public string Code { get; set; }

        public string Type { get; set; }

        public string  Value {get;set;}
    }
}
