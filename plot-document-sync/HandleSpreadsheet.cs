using System.Text.RegularExpressions;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace plot_document_sync;

public static class HandleSpreadsheet
{
    public const string SpreadsheetId = "1ORdiQiXm-SsRq1xVmK-vZBsyYZ1SE2FeihafNfR5N-g";
    public const string SheetRange = "Pilot";
    public const bool OverwriteValues = false;

    private static readonly string[] Characters =
    {
        "ANDY",
        "LIVI",
        "ENZO"
    };
    
    private static int _dialogueIndex = -1;
    private static int _cameraIndex = -1;
    private static int _expressionIndex = -1;
    
    private static void SetHeaderIndexes(ValueRange spreadsheet)
    {
        for (int columnIndex = 0; columnIndex < spreadsheet.Values[0].Count; columnIndex++)
        {
            string cell = spreadsheet.Values[0][columnIndex].ToString() ?? string.Empty;

            switch (cell.ToLower())
            {
                case "dialogue":
                    _dialogueIndex = columnIndex;
                    continue;
                case "camera":
                    _cameraIndex = columnIndex;
                    continue;
                case "expression":
                    _expressionIndex = columnIndex;
                    continue;
            }
        }

        if (_dialogueIndex == -1) throw new Exception("Unable to find dialogue header");
        if (_cameraIndex == -1) throw new Exception("Unable to find camera header");
        if (_expressionIndex == -1) throw new Exception("Unable to find expression header");
        
        // Console.WriteLine($"Dialogue Index: {_dialogueIndex}");
        // Console.WriteLine($"Camera Index: {_cameraIndex}");
        // Console.WriteLine($"Expression Index: {_expressionIndex}");
    }

    public static ValueRange EditSheetData(ValueRange spreadsheet, IReadOnlyList<string> documentContent)
    {
        SetHeaderIndexes(spreadsheet);

        spreadsheet.Values[0] = new object[spreadsheet.Values[0].Count];
        
        for (int rowIndex = 1; rowIndex <= documentContent.Count; rowIndex++)
        {
            object[] row = new object[spreadsheet.Values[0].Count];
            try
            {
                var _ = spreadsheet.Values[rowIndex];
            }
            catch
            {
                spreadsheet.Values.Add(row);
            }
            string documentParagraph = documentContent[rowIndex - 1];

            // Dialogue
            try
            {
                if (documentParagraph != spreadsheet.Values[rowIndex][_dialogueIndex].ToString())
                {
                    row[_dialogueIndex] = documentParagraph;
                }
            }
            catch
            {
                row[_dialogueIndex] = documentParagraph;
            }


            // Camera
            foreach (string character in Characters)
            {
                try
                {
                    if (!documentParagraph.StartsWith(character)
                        || spreadsheet.Values[rowIndex][_cameraIndex] is not null
                        && !OverwriteValues) continue;
                }
                catch
                {
                    // ignored
                }
                
                row[_cameraIndex] = $"PlayerCam looking at {character}";
                break;
            }
            
            // Expression
            string? expression = null;
            Regex expressionRx = new(@"\((.*?)\)", RegexOptions.Compiled);
            MatchCollection expressionMatches = expressionRx.Matches(documentParagraph);
            if (expressionMatches.Any())
            {
                expression = expressionMatches[0].Value;
            }

            try
            {
                if (spreadsheet.Values[rowIndex][_cameraIndex] is null || OverwriteValues)
                {
                    row[_expressionIndex] = expression;
                }
            }
            catch
            {
                row[_expressionIndex] = expression;
            }

            spreadsheet.Values[rowIndex] = row.ToList();
        }

        return spreadsheet;
    }

    public static void WriteSheetData(ValueRange spreadsheet, SpreadsheetsResource.ValuesResource spreadsheetResource)
    {
        ValueRange updateValues = new()
        {
            Values = spreadsheet.Values.Take(Program.SheetRows).ToList()
        };

        ValueRange appendValues = new()
        {
            Values = spreadsheet.Values.Skip(Program.SheetRows).ToList()
        };

        var updateRequest = spreadsheetResource.Update(updateValues, SpreadsheetId, SheetRange);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        UpdateValuesResponse updateRequestResult = updateRequest.Execute();

        var appendRequest =
            spreadsheetResource.Append(appendValues, SpreadsheetId, $"{SheetRange}!A{Program.SheetRows + 1}:L");
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
        AppendValuesResponse appendValuesResult = appendRequest.Execute();
        
        
        Console.WriteLine($"{updateRequestResult.UpdatedCells ?? 0} cells updated");
        Console.WriteLine($"{appendValuesResult.Updates.UpdatedRows ?? 0} rows appended");
    }
}