using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace Config
{
    class RewardParameters
    {
        public int count { get; set; }
        public string item_id { get; set; }
        public double chance { get; set; }
    }

    class RewardConfig
    {
        public string type { get; set; }
        public RewardParameters parameters { get; set; }
    }

    class GroupConfig
    {
        public double group_chance { get; set; }
        public Dictionary<string, RewardConfig> rewards { get; set; }

        public GroupConfig(double chance)
        {
            group_chance = chance;
            rewards = new Dictionary<string, RewardConfig>();
        }
    }

    class CaseConfig
    {
        public Dictionary<string, GroupConfig> groups { get; set; }

        public CaseConfig()
        {
            groups = new Dictionary<string, GroupConfig>();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var baseDir = @"C:\Users\Korov\Documents\ChillBase\";
            var excelPath = Path.Combine(baseDir, "Tech GD Test Config.xlsx");
            var jsonPath = Path.Combine(baseDir, "case_config.json");

            var wb = new XLWorkbook(excelPath);
            var casesSheet = wb.Worksheet("Кейсы Парсинг");
            var rewardsSheet = wb.Worksheet("Таблица Наград");

            var currencyMap = new Dictionary<string, (string naming, string id, string type)>();
            foreach (var r in rewardsSheet.RowsUsed().Skip(1))
            {
                var name = r.Cell(1).GetString().Trim();
                var techNaming = r.Cell(2).GetString().Trim();
                var itemId = r.Cell(3).GetString().Trim().Replace(" ", "");
                var type = r.Cell(4).GetString().Trim();
                if (!string.IsNullOrEmpty(name))
                    currencyMap[name] = (techNaming, itemId, type);
            }
            var rewardsMap = new Dictionary<string, (string naming, string id, string type)>();
            foreach (var r in rewardsSheet.RowsUsed().Skip(2))
            {
                var name = r.Cell(6).GetString().Trim();
                var techNaming = r.Cell(7).GetString().Trim();
                var itemId = r.Cell(8).GetString().Trim().Replace(" ", "");
                var type = r.Cell(9).GetString().Trim();
                if (!string.IsNullOrEmpty(name))
                    rewardsMap[name] = (techNaming, itemId, type);
            }

            var config = new Dictionary<string, CaseConfig>();

            string lastCase = null;
            string lastGroup = null;
            double lastGChance = 0;

            foreach (var row in casesSheet.RowsUsed().Skip(1))
            {
                var cellCase = row.Cell(1).GetString().Trim();
                if (!string.IsNullOrEmpty(cellCase))
                    lastCase = cellCase;                  
                                                          
                var caseKey = lastCase;

                var cellGroup = row.Cell(2).GetString().Trim();
                if (!string.IsNullOrEmpty(cellGroup))
                {
                    lastGroup = cellGroup;
                    lastGChance = ParsePercent(row.Cell(3).GetString());
                }
                var groupKey = lastGroup;
                var groupChance = lastGChance;

                var rewardName = row.Cell(4).GetString().Trim();
                var rewardCount = row.Cell(5).GetValue<int>();
                var rewardChance = ParsePercent(row.Cell(6).GetString());

                if (!config.ContainsKey(caseKey))
                    config[caseKey] = new CaseConfig();
                var caseCfg = config[caseKey];

                if (!caseCfg.groups.ContainsKey(groupKey))
                    caseCfg.groups[groupKey] = new GroupConfig(groupChance);
                var groupCfg = caseCfg.groups[groupKey];

                (string naming, string id, string type) info;
                if (!currencyMap.TryGetValue(rewardName, out info))
                    rewardsMap.TryGetValue(rewardName, out info);

                groupCfg.rewards[info.naming] = new RewardConfig
                {
                    type = info.type,
                    parameters = new RewardParameters
                    {
                        count = rewardCount,
                        item_id = info.id,
                        chance = rewardChance
                    }
                };
            }

            var json = JsonConvert.SerializeObject(config, Formatting.Indented);
            File.WriteAllText(jsonPath, json, System.Text.Encoding.UTF8);
            Console.WriteLine("Конфиг записан в " + jsonPath);
        }
        static double ParsePercent(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
                return 0.0;

            var s = rawValue.Trim();

            var isPercent = s.EndsWith("%", StringComparison.Ordinal);

            s = s.TrimEnd('%')
                 .Replace(',', '.');

            if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
                return 0.0;

            return isPercent
                ? value / 100.0
                : value;
        }
    }
}
