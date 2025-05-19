using ClosedXML.Excel;
using System.Text;
using System.Text.RegularExpressions;
class StepInfo
{
    public string StepName { get; set; }
    public string ElementName { get; set; }
    public string XPath { get; set; }
    public string ParameterName { get; set; }
    public string ParameterValue { get; set; }
}
class Program
{
    public static string featureName;
    public static string Namespace;
    static void Main()
    {
        Console.Write("Enter feature name (e.g., Login): ");
        featureName = Console.ReadLine().Trim();

        Console.Write("Enter Namespace: ");
        Namespace = Console.ReadLine().Trim();

        Console.Write("Enter path to Excel file (e.g., C:\\Steps.xlsx): ");
        string excelPath = Console.ReadLine().Trim();

        var steps = ReadStepsFromExcel(excelPath);

        if (!Directory.Exists(featureName))
            Directory.CreateDirectory(featureName);

        GenerateFeatureFile(featureName, steps);
        GenerateElementsFile(featureName, steps);
        GeneratePageFile(featureName, steps);
        GenerateStepsFile(featureName, steps);

        Console.WriteLine($"\n✅ Code generated in ./{featureName}/ folder.");

        string outputPath = Path.Combine(featureName, $"{featureName}.xlsx");
        SaveStepsToExcel(outputPath, steps);

        Console.WriteLine($"\nExcel file generated successfully at: {outputPath}");
    }
    static List<StepInfo> ReadStepsFromExcel(string filePath)
    {
        var steps = new List<StepInfo>();
        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheets.First();

        var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

        foreach (var row in rows)
        {
            string stepName = row.Cell(1).GetString().Trim();
            string xpath = row.Cell(2).GetString().Trim();
            string elementName = ExtractElementNameFromXPath(xpath);
            string parameterName = ExtractParameterNameFromStep(stepName);
            var quotedMatch = Regex.Match(stepName, "\"(?<quoted>[^\"]+)\"");
            string parameterValue = quotedMatch.Success ? quotedMatch.Groups["quoted"].Value : "";

            steps.Add(new StepInfo
            {
                StepName = stepName,
                XPath = xpath,
                ElementName = elementName,
                ParameterName = parameterName,
                ParameterValue = parameterValue
            });
        }

        return steps;
    }
    static string ExtractParameterNameFromStep(string step)
    {
        var angleMatch = Regex.Match(step, @"<(?<param>[^>]+)>");
        if (angleMatch.Success)
            return angleMatch.Groups["param"].Value;

        var quotedMatch = Regex.Match(step, "\"(?<quoted>[^\"]+)\"");
        if (quotedMatch.Success)
            return "Param";

        return string.Empty;
    }
    static void SaveStepsToExcel(string outputPath, List<StepInfo> steps)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add($"{featureName}");

        var elementNames = steps
            .Where(step => !string.IsNullOrEmpty(step.ParameterValue))
            .Select(step => step.ElementName)
            .Distinct()
            .ToList();

        int column = 1;

        foreach (var elementName in elementNames)
        {
            ws.Cell(1, column++).Value = elementName;
        }

        column = 1;
        foreach (var elementName in elementNames)
        {
            var matchingStep = steps.FirstOrDefault(step => step.ElementName == elementName && !string.IsNullOrEmpty(step.ParameterValue));

            if (matchingStep != null)
            {
                ws.Cell(2, column).Value = matchingStep.ParameterValue;
            }

            column++;
        }

        workbook.SaveAs(outputPath);
    }

    static string ExtractElementNameFromXPath(string xpath)
    {
        var nameMatch = Regex.Match(xpath, @"@name='([^']+)'");
        if (nameMatch.Success)
            return ToPascalCase(nameMatch.Groups[1].Value);

        var idMatch = Regex.Match(xpath, @"@id='([^']+)'");
        if (idMatch.Success)
            return ToPascalCase(idMatch.Groups[1].Value);

        var classMatch = Regex.Match(xpath, @"@class='([^']+)'");
        if (classMatch.Success)
            return ToPascalCase(classMatch.Groups[1].Value.Split(' ')[0]);

        var normalizeSpaceMatch = Regex.Match(xpath, @"normalize-space\(\)\s*=\s*'([^']+)'");
        if (normalizeSpaceMatch.Success)
            return ToPascalCase(normalizeSpaceMatch.Groups[1].Value);

        var textMatch = Regex.Match(xpath, @"text\(\)\s*=\s*'([^']+)'");
        if (textMatch.Success)
            return ToPascalCase(textMatch.Groups[1].Value);

        var tagTypeMatch = Regex.Match(xpath, @"//(\w+)\[@type='([^']+)'\]");
        if (tagTypeMatch.Success)
            return ToPascalCase(tagTypeMatch.Groups[2].Value);

        var tagMatch = Regex.Match(xpath, @"//(\w+)");
        if (tagMatch.Success)
        {
            var tag = tagMatch.Groups[1].Value;
            var attrMatch = Regex.Match(xpath, @"@(\w+)='([^']+)'");
            if (attrMatch.Success)
            {
                var attrValue = attrMatch.Groups[2].Value;
                return ToPascalCase(tag + "_" + attrValue.Replace(" ", "").Replace("/", ""));
            }
            return ToPascalCase(tag);
        }

        return "unknownElement";
    }

    static string ToPascalCase(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return input;
        return string.Concat(input
            .Split(new[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => char.ToUpperInvariant(s[0]) + s.Substring(1)));
    }

    static string DetectActionFromStepName(string stepName)
    {
        stepName = stepName.ToLower();
        if (stepName.Contains("click")) return "click";
        if (stepName.Contains("enter") || stepName.Contains("sendkeys")) return "sendkeys";
        if (stepName.Contains("select")) return "select";
        if (stepName.Contains("verify")) return "verify";
        return "default";
    }

    static void GenerateFeatureFile(string feature, List<StepInfo> steps)
    {
        var sb = new StringBuilder();

        string readableFeature = Regex.Replace(feature, "(\\B[A-Z])", " $1");

        sb.AppendLine($"Feature: {feature}");
        sb.AppendLine($"@{feature}");
        sb.AppendLine($"@DataSource:../../../../Teacher/Excel/{featureName}.xlsx @DataSet:{featureName}");
        sb.AppendLine($"  Scenario: {readableFeature} Test");

        foreach (var step in steps)
        {
            string detectedAction = DetectActionFromStepName(step.StepName);

            string keyword = detectedAction switch
            {
                "click" => "When",
                "sendkeys" => "When",
                "select" => "When",
                "verify" => "Then",
                _ => "Given"
            };

            string updatedStepName = step.StepName;

            var matches = Regex.Matches(updatedStepName, "\"(.*?)\"");
            foreach (Match match in matches)
            {
                string dynamicValue = match.Groups[1].Value;
                string placeholder = $"\"<{step.ElementName}>\"";
                updatedStepName = updatedStepName.Replace(match.Value, placeholder);
            }

            updatedStepName = updatedStepName.Replace("<>", step.ElementName);

            sb.AppendLine($"    {keyword} {readableFeature} {updatedStepName}");
        }

        Directory.CreateDirectory(feature);
        File.WriteAllText($"{feature}/{feature}.feature", sb.ToString());
    }

    static void GenerateElementsFile(string feature, List<StepInfo> steps)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("using OpenQA.Selenium;");
        sb.AppendLine();
        sb.AppendLine($"public class {feature}Elements");
        sb.AppendLine("{");

        foreach (var step in steps)
        {
            sb.AppendLine($"    public static By {step.ElementName} => By.XPath(\"{step.XPath}\");");
        }

        sb.AppendLine("}");
        File.WriteAllText($"{feature}/{feature}Element.cs", sb.ToString());
    }

    static void GeneratePageFile(string feature, List<StepInfo> steps)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("using OpenQA.Selenium;");
        sb.AppendLine();
        sb.AppendLine($"public class {feature}Page(IWebDriver driver)");
        sb.AppendLine("{");

        foreach (var step in steps)
        {
            sb.AppendLine($"    public IWebElement Get{step.ElementName}() => driver.FindElement({feature}Elements.{step.ElementName});");
        }

        sb.AppendLine("}");
        File.WriteAllText($"{feature}/{feature}Page.cs", sb.ToString());
    }

    static void GenerateStepsFile(string feature, List<StepInfo> steps)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"namespace UMS.UI.Test.ERP.Areas.{Namespace}.Steps");
        sb.AppendLine("{");
        sb.AppendLine();

        string readableFeature = Regex.Replace(feature, "([a-z])([A-Z])", "$1 $2");

        sb.AppendLine("[Binding]");
        sb.AppendLine($"public class {feature}Step({feature}Page page)");
        sb.AppendLine("{");
        sb.AppendLine();

        var elementLookup = steps
            .GroupBy(s => s.ElementName)
            .ToDictionary(g => g.Key, g => g.First());

        foreach (var step in steps)
        {
            string stepText = step.StepName;
            var paramNames = new List<string>();

            var matchesAngle = Regex.Matches(stepText, "<(?<param>.*?)>");
            paramNames.AddRange(matchesAngle.Cast<Match>().Select(m => m.Groups["param"].Value));

            var matchesQuoted = Regex.Matches(stepText, "\"(?<quoted>[^\"]+)\"");
            foreach (Match match in matchesQuoted)
            {
                if (!string.IsNullOrEmpty(step.ElementName))
                {
                    string clean = Regex.Replace(step.ElementName, @"[^a-zA-Z0-9]", "");
                    string camel = char.ToLower(clean[0]) + clean.Substring(1);
                    if (!paramNames.Contains(camel))
                        paramNames.Add(camel);
                }
                else
                {
                    paramNames.Add("param" + (paramNames.Count + 1));
                }
            }

            paramNames = paramNames
                .Select(p => Regex.Replace(p, @"[^a-zA-Z0-9]", ""))
                .Distinct()
                .ToList();

            string specFlowStepText = Regex.Replace(stepText, "<.*?>", "{string}");
            specFlowStepText = Regex.Replace(specFlowStepText, "\"[^\"]+\"", "{string}");

            string paramList = string.Join(", ", paramNames.Select(p =>
            {
                var cleaned = Regex.Replace(p, @"[^a-zA-Z0-9]", "");
                return $"string {char.ToLower(cleaned[0]) + cleaned.Substring(1)}";
            }));

            string cleanStepName = Regex.Replace(step.StepName, "<.*?>", "");
            cleanStepName = Regex.Replace(cleanStepName, "\".*?\"", "");
            string methodName = $"{feature}{Regex.Replace(cleanStepName, "[^a-zA-Z0-9]", "")}";

            var elementStep = elementLookup[step.ElementName];
            string elementCall = $"page.Get{elementStep.ElementName}()";

            string detectedAction = DetectActionFromStepName(step.StepName);
            string actionLine = DetermineAction(detectedAction, paramNames, elementCall);

            string keyword = detectedAction switch
            {
                "click" => "When",
                "sendkeys" => "When",
                "select" => "When",
                "verify" => "Then",
                _ => "Given"
            };

            sb.AppendLine($"    [{keyword}(\"{readableFeature} {specFlowStepText}\")]");
            sb.AppendLine($"    public void {keyword}{methodName}({paramList})");
            sb.AppendLine("    {");

            if (detectedAction == "sendkeys")
            {
                actionLine = $"{elementCall}.SendKeys({paramNames[0]});";
            }

            sb.AppendLine($"        {actionLine}");
            sb.AppendLine("    }");
            sb.AppendLine();
        }

        sb.AppendLine("  }");
        sb.AppendLine("}");
        Directory.CreateDirectory(feature);
        File.WriteAllText($"{feature}/{feature}Step.cs", sb.ToString());
    }

    static string DetermineAction(string detectedAction, List<string> paramNames, string elementCall)
    {
        string param = (paramNames.Count > 0) ? FormatParam(paramNames[0]) : "\"value\"";

        return detectedAction switch
        {
            "click" => $"{elementCall}.Click();",
            "sendkeys" => $"{elementCall}.SendKeys({param});",
            "select" => $"new SelectElement({elementCall}).SelectByText({param});",
            "verify" => $"// Add your verification logic for {elementCall}",
            _ => "// Unsupported action"
        };
    }

    static string FormatParam(string param)
    {
        var cleaned = Regex.Replace(param, @"[^a-zA-Z0-9]", "");
        if (string.IsNullOrEmpty(cleaned)) return "param";
        cleaned = char.ToLower(cleaned[0]) + cleaned.Substring(1);
        return char.IsLetter(cleaned[0]) ? cleaned : $"_{cleaned}";
    }
}
