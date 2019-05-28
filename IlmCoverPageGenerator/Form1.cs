using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IlmCoverPageGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // prepare cover data
            var prepareCover = true;
            prepareCoverData(prepareCover);
            // get cover data
            // create covers
            // prepare module
            // append covers
            // review covers
        }

        private void prepareCoverData(bool v)
        {
            if (v == true)
            {
                List<ModuleInfo> allModulesInfo = new List<ModuleInfo>();
                var allData = File.ReadAllLines(@"C:\Users\kstaples\Documents\Projects\ILM Script\ILM\ILM_ModuleList_ExtractTradeSecrets.csv");
                for (var i = 0; i < allData.Length; i++)
                {
                    var recordData = allData[i].Split(',');
                    var moduleNumber = allData[0];
                    var modulePeriod = allData[1];
                    var moduleTrade = allData[2];
                    var moduleTitle = allData[3].Replace("[comma]", ",");

                    var module = new ModuleInfo(moduleNumber, modulePeriod, moduleTrade, moduleTitle);
                    allModulesInfo.Add(module);
                }
                var outData = JsonConvert.SerializeObject(allModulesInfo);

                var outFilePath = @"C:\Users\kstaples\Documents\Projects\ILM Covers\CoverData.json";
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(outFilePath))
                {
                    file.Write(outData);
                }
            }
            else { };
        }

        [Serializable]
        public class ModuleInfo
        {
            public string ModuleTrade { get; set; }
            public string ModuleNumber { get; set; }
            public string ModuleTitle { get; set; }
            public string ModulePeriod { get; set; }

            public ModuleInfo(string moduleNumber, string modulePeriod, string moduleTitle, string moduleTrade)
            {
                ModuleNumber = moduleNumber;
                ModulePeriod = ToPeriodString(modulePeriod);
                ModuleTitle = moduleTitle;
                ModuleTrade = moduleTrade;
            }

            private string ToPeriodString(string modulePeriod)
            {
                int periodInt = Convert.ToInt32(modulePeriod.Split('_')[1]);
                if (periodInt == 1)
                {
                    return "FIRST PERIOD";
                }
                else if (periodInt == 2)
                {
                    return "SECOND PERIOD";
                }
                else if (periodInt == 3)
                {
                    return "THIRD PERIOD";
                }
                else if (periodInt == 4)
                {
                    return "FOURTH PERIOD";
                }
                else if (periodInt == 0)
                {
                    return "ZERO PERIOD";
                }
                return "";
            }
        }
    }
}
