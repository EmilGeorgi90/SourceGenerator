using System.Collections.Immutable;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.Office.Interop.Excel;

namespace FunctionsGenerator;

[Generator]
public class FunctionGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var provider = context.SyntaxProvider.CreateSyntaxProvider(
            predicate: (c, _) => c is ClassDeclarationSyntax,
            transform: (n, _) => (ClassDeclarationSyntax)n.Node
        ).Where(m => m is not null);
        var compilation = context.CompilationProvider.Combine(provider.Collect());
        context.RegisterSourceOutput(compilation, (spc, source) => Execute(spc, source.Left, source.Right));
    }
    private void Execute(SourceProductionContext spc, Compilation compilation,
        ImmutableArray<ClassDeclarationSyntax> typeList)
    {
        #region vars
        var bldr = new StringBuilder();
        var classes = new StringBuilder();
        var methods = new StringBuilder();
        var enums = new StringBuilder();
        Microsoft.Office.Interop.Excel.Application excel = new();
        Workbook wb = excel.Workbooks.Open(@"C:\Users\emil\RiderProjects\SourceGenerator\FunctionsGenerator\ImpactMapOptimizedForSROIReportingEmil.xlsx");
        #endregion
        try
        {
            #region start of 
            var excelDatas = ExcelData(excel, wb);
            bldr.AppendLine("using App;");
            bldr.AppendLine("namespace SampleSourceGenerator");
            bldr.AppendLine("{");
            classes.AppendLine("""  public static partial class ClassNames""");
            classes.AppendLine("""  {""");
            #endregion
            #region generate methods
            foreach (var data in excelDatas)
            {
                methods.AppendLine();
                switch (data.Action)
                {
                    case Action.SUM:
                        try
                        {
                            var cells = data.Formula.Split('(')[1];
                            var cell1 = cells.Split('+')[0];
                            var cell2 = cells.Split('+')[1];
                            cell2 = cell2.Remove(cell2.Length - 1, 1);
                            var methodName = GetExcelText(excel, wb, $"{ToLetter(data.Column)}{data.Row}");
                            methodName = string.IsNullOrEmpty(methodName) ? $"{ToLetter(data.Column)}{data.Row}" : methodName;
                            var param1 = GetExcelText(excel, wb, cell1);
                            var param2 = GetExcelText(excel, wb, cell2);
                            param1 = string.IsNullOrWhiteSpace(param1) ? "param1" : param1.Split(' ')[0];
                            param2 = string.IsNullOrWhiteSpace(param2) ? "param1" : param2.Split(' ')[0];
                            methods.AppendLine($"""
                                /// <summary>
                                /// {data.Text}
                                /// </summary>
                                /// <param name="{param1}">{GetExcelCellData(excel, wb, $"{cell1}9")}</param>
                                /// <param name="{param2}">{GetExcelCellData(excel, wb, $"{cell2}9")}</param>
                                /// <returns>{param1} + {param2}</returns>
                                """);
                            methods.AppendLine($""" public static int {methodName}(int {param1}, int {param2}) => {param1} + {param2};""");
                            break;
                        }
                        catch (Exception e)
                        {
                            methods.AppendLine($"//{e.Message}");
                            break;
                        }
                    case Action.IF:
                        var param = data.Formula.Split('(')[1].Split(',');
                        var cell = param[0].Split('=')[0];
                        var cellText = GetExcelText(excel, wb, cell).Trim();
                        cellText = string.IsNullOrWhiteSpace(cellText) ? "param1" : cellText;
                        cellText = Regex.Replace(cellText, @"\b\w", m => m.Value.ToUpper()).Replace(" ", "");
                        var segments = SplitOnSpecialCharactersAndParentheses(data.Formula);
                        var statement = param[0].Split('=').Last();
                        statement = Regex.Replace(statement, @"\b\w", m => m.Value.ToUpper()).Replace(" ", "");
                        var ifTrue = param[1];
                        ifTrue = Regex.Replace(ifTrue, @"\b\w", m => m.Value.ToUpper()).Replace(" ", "");
                        var ifFalse = param[2];
                        ifFalse = Regex.Replace(ifFalse, @"\b\w", m => m.Value.ToUpper()).Replace(" ", "");
                        var name = string.IsNullOrWhiteSpace(data.Text) ? $"{ToLetter(data.Column)}{data.Row}" : int.TryParse(data.Text.Substring(0, 1), out int _) ? $"{ToLetter(data.Column)}{data.Row}" : data.Text;
                        name = Regex.Replace(name, @"\b\w", m => m.Value.ToUpper()).Replace(" ", "");
                        var isNumeric = int.TryParse(segments.Text[1], out var _);
                        string type = isNumeric ? "int" : "string";
                        var operation = segments.Operation[0] == "=" ? "==" : segments.Operation[0];
                        methods.AppendLine($"""
                                            //{string.Join(",", segments.Text[1])}
                                            //{string.Join(",", operation)}
                                            public static int {name}({type} {cellText}) => ({cellText} {operation} {segments.Text[1].Split(',')[0]} ? {ifTrue} : {ifFalse};
                                            """);
                        break;
                    case Action.SUMIF:

                        methods.AppendLine($""" //formula is {data.Formula} """);
                        break;
                    case Action.UNKNOWN:
                        methods.AppendLine($""" //{data.Formula} has an unknown action  """);
                        break;
                    case Action.DropDown:
                        var elements = data.Formula.Split(',');
                        if (int.TryParse(elements[0], out int _))
                        {
                            break;
                        }
                        enums.AppendLine($$"""
                            public enum {{data.Text}}
                            {
                            """);
                        for (var i = 0; i < elements.Length; i++)
                        {
                            var element = Regex.Replace(elements[i], @"\b\w", m => m.Value.ToUpper()).Replace(" ", "");
                            enums.AppendLine(element + ",");
                        }
                        enums.AppendLine($$"""
                            }
                            """);
                        break;
                    default:
                        break;
                }
                methods.AppendLine();
            }
            #endregion
            #region end of class
            classes.AppendLine($"""        {methods}""");
            classes.AppendLine("""  }""");
            bldr.AppendLine(classes.ToString());
            bldr.AppendLine("}");
            bldr.AppendLine(enums.ToString());
            spc.AddSource("ClassNames.g.cs", bldr.ToString());
            #endregion
        }
        catch (Exception ex)
        {
            bldr.AppendLine($"//error: {ex.Message}");
            throw;
        }
        finally
        {
            wb.Close(false);
            excel.Quit();
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excel);
        }
    }
    /// <summary>
    /// split of Special characters used to split formulars from excel to get math and other function from cell
    /// </summary>
    /// <param name="input">cell text or cell formular</param>
    /// <returns>the input splitted</returns>

    private static Formula SplitOnSpecialCharactersAndParentheses(string input)
    {
        List<string> segments = new List<string>();
        List<string> operators = new List<string>();
        StringBuilder currentSegment = new StringBuilder();
        Stack<char> parenthesesStack = new Stack<char>();
        bool foundFirstParenthesis = false;

        foreach (char c in input)
        {
            if (c == '(')
            {
                if (currentSegment.Length > 0)
                {
                    if (!foundFirstParenthesis)
                    {
                        // Clear segments if we haven't found the first parenthesis yet
                        segments.Clear();
                        foundFirstParenthesis = true;
                    }
                    else
                    {
                        segments.Add(currentSegment.ToString());
                    }
                    currentSegment.Clear();
                }
                parenthesesStack.Push(c);
            }
            else if (c == ')')
            {
                if (parenthesesStack.Count > 0)
                {
                    currentSegment.Append(c);
                    if (parenthesesStack.Peek() == '(')
                    {
                        parenthesesStack.Pop();
                        string segmentWithoutParentheses = currentSegment.ToString().Trim('(', ')');
                        segments.Add(segmentWithoutParentheses);
                        currentSegment.Clear();
                    }
                }
            }
            else if (IsMathOperator(c.ToString()))
            {
                if (currentSegment.Length > 0)
                {
                    segments.Add(currentSegment.ToString());
                    currentSegment.Clear();
                }
                operators.Add(c.ToString());
            }
            else
            {
                currentSegment.Append(c);
            }
        }

        // If there's a remaining non-matched substring, add it to segments
        if (currentSegment.Length > 0)
        {
            segments.Add(currentSegment.ToString());
        }

        // Create the Formula struct
        Formula formula = new Formula
        {
            Text = segments.ToArray(),
            Operation = operators.ToArray()
        };

        // Remove parentheses from segments
        for (int i = 0; i < formula.Text.Length; i++)
        {
            formula.Text[i] = formula.Text[i].Trim('(', ')');
        }

        return formula;
    }
    private static bool IsMathOperator(string segment)
    {
        // Add more mathematical operators as needed
        string[] mathOperators = { "+", "-", "*", "/", "=", ">", "<", ">=", "<=", "==", "!=" };
        return Array.Exists(mathOperators, op => op.Equals(segment));
    }

    /// <summary>
    /// Get excel text from Column and row 7, used to get method names, STILL WORK IN PROGRESS
    /// </summary>
    /// <param name="excel">excel app</param>
    /// <param name="wb">current Workbook</param>
    /// <param name="cell">only the column in letter notation</param>
    /// <returns></returns>
    private static string GetExcelText(Application excel, Workbook wb, string cell)
    {
        try
        {
            string result = string.Empty;
            string textPattern = @"[A-Z]+";
            var columnText = Regex.Match(cell, textPattern, RegexOptions.IgnoreCase);
            if (columnText.Success)
            {
                var column = ToNumber(columnText.Value);
                foreach (Worksheet sheet in wb.Sheets)
                {
                    dynamic c = sheet.Cells[7, column];
                    result = c.Text;
                }
            }
            return result;

        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            return string.Empty;
        }
    }
    /// <summary>
    /// Get excel cell data from specific Cell
    /// </summary>
    /// <param name="excel">excel app</param>
    /// <param name="wb">current Workbook</param>
    /// <param name="cell">the cell you want data from in the excel format f.eks A2</param>
    /// <returns>Text from cell provided</returns>
    private static string GetExcelCellData(Application excel, Workbook wb, string cell)
    {
        string result = string.Empty;
        string textPattern = @"[A-Z]+";
        string numberPattern = @"[0-9]+";
        var columnText = Regex.Match(cell, textPattern, RegexOptions.IgnoreCase).Value;
        var row = Regex.Match(cell, numberPattern, RegexOptions.IgnoreCase).Value;
        var column = ToNumber(columnText);
        foreach (Worksheet sheet in wb.Sheets)
        {
            dynamic c = sheet.Cells[row, column];
            result = c.Text;
        }
        return result;
    }
    /// <summary>
    /// convert column from number notation to letter notation
    /// </summary>
    /// <param name="columnNumber"></param>
    /// <returns>Column name as letter notation</returns>
    /// <exception cref="ArgumentOutOfRangeException">Column number must be greater then or equal to 1</exception>
    private static string ToLetter(int columnNumber)
    {
        if (columnNumber < 1)
        {
            throw new ArgumentOutOfRangeException("Column number must be greater than or equal to 1.");
        }

        string result = string.Empty;

        while (columnNumber > 0)
        {
            int remainder = (columnNumber - 1) % 26;
            char digit = (char)('A' + remainder);

            result = digit + result;
            columnNumber = (columnNumber - 1) / 26;
        }

        return result;
    }

    /// <summary>
    /// convert column name from letter notation to number notation
    /// </summary>
    /// <param name="columnName"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException"></exception>
    private static int ToNumber(string columnName)
    {
        if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");
        columnName = columnName.ToUpperInvariant();
        int sum = 0;
        for (int i = 0; i < columnName.Length; i++)
        {
            sum *= 26;
            sum += (columnName[i] - 'A' + 1);
        }
        return sum;
    }

    /// <summary>
    /// Used to extract data and dropdown items
    /// </summary>
    /// <param name="excel">excel app</param>
    /// <param name="wb">current Workbook</param>
    /// <returns>Extracted data</returns>
    private static List<ExcelDataType> ExcelData(Application excel, Workbook wb)
    {
        var result = new List<ExcelDataType>();
        string exceptions = string.Empty;

        foreach (Worksheet sheet in wb.Sheets)
        {
            foreach (Range column in sheet.UsedRange.Columns)
            {
                try
                {
                    Range cell = (Range)sheet.Cells[10, column.Column];
                    if (cell != null)
                    {
                        try
                        {
                            List<string> dropdownOptions = new();
                            #region dropdown into enum
                            if (cell.Validation.Type == (int)XlDVType.xlValidateList)
                            {
                                Validation validation = cell.Validation;
                                string formula1 = validation.Formula1;

                                if (formula1.StartsWith("="))
                                {
                                    Range range = sheet.Evaluate(formula1) as Range;

                                    if (range != null)
                                    {
                                        foreach (Range valueCell in range)
                                        {
                                            dropdownOptions.Add(valueCell.Value2.ToString());
                                        }
                                    }
                                }
                                else
                                {
                                    string[] values = formula1.Split(',');

                                    foreach (string value in values)
                                    {
                                        dropdownOptions.Add(value.Trim());
                                    }
                                }
                                result.Add(new ExcelDataType()
                                { Formula = string.Join(",", dropdownOptions).ToString(), Row = 10, Column = column.Column, Text = $"{ToLetter(column.Column)}10", Action = Action.DropDown });
                            }
                            #endregion
                        }
                        catch (COMException ex)
                        {
                            exceptions += ex.Message;
                        }
                        catch (Exception ex)
                        {
                            throw;
                        }
                        #region extract formula
                        if (((string)cell.Formula).StartsWith("="))
                        {
                            var formula = ((string)cell.Formula).Remove(0, 1);
                            var text = (string)((Range)sheet.Cells[7, column.Column]).Text;
                            Action action = Action.UNKNOWN;
                            int openParenIndex = formula.IndexOf('(');
                            if (openParenIndex >= 0)
                            {
                                string extractedText = formula.Substring(0, openParenIndex);
                                if (Enum.TryParse(extractedText, out Action a))
                                {
                                    action = a;
                                }
                            }
                            result.Add(new ExcelDataType()
                            { Formula = formula, Row = 10, Column = column.Column, Text = text, Action = action });
                        }
                        #endregion
                    }
                }
                catch (Exception e)
                {
                    exceptions += e.Message;
                    //throw;
                }

            }

        }
        if (exceptions.Length > 0)
        {
            result.Add(new ExcelDataType()
            {
                Action = Action.UNKNOWN,
                Column = 0,
                Row = 10,
                Formula = exceptions,
            });
        }
        return result;
    }

}
/// <summary>
/// This struct is a representation of the extracted excel data
/// </summary>
public struct ExcelDataType
{
    public string Text { get; set; }
    public int Row { get; set; }
    public int Column { get; set; }

    public string Formula { get; set; }
    public Action Action { get; set; }
}

/// <summary>
/// current know action from excel, PLEASE EXTEND ME
/// </summary>
public enum Action
{
    SUM,
    IF,
    SUMIF,
    DropDown,
    UNKNOWN
}
public struct Formula
{
    public string[] Text { get; set; }
    public string[] Operation { get; set; }
}