using System;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

public class Countries {
	public Name name { get; set; }
	public double? area { get; set; }
	public List<string>? capital { get; set; }
	public Dictionary<string, Currencies> currencies { get; set; }
	public List<string>? currencyCodes { get; set; }
}

public class Name {
	public string common { get; set; }
}

public class Currencies {
	public string? symbol;
	public string? name;
}

public class Program{
	static readonly HttpClient client = new HttpClient();
	static List<Countries> listOfCountries;	// List to store countries' information
	
	static async Task Main(){
		await FetchInfo();

		await MakeExcel();
	}

	static async Task FetchInfo() {
		string jsonUrl = @"https://restcountries.com/v3.1/all?fields=name,capital,area,currencies";
		string jsonString = "";

	// FETCH JSON STRING
		try{
			jsonString = await client.GetStringAsync(jsonUrl);
		}
		catch(HttpRequestException e){
			jsonString = "";
			Console.WriteLine("\nException Caught!");	
			Console.WriteLine("Message :{0} ",e.Message);

			return;
		}

	// SORT LIST ALPHABETICALLY
		listOfCountries = JsonConvert.DeserializeObject<List<Countries>>(jsonString);
		listOfCountries.Sort(delegate(Countries x, Countries y){
            if (x.name.common == null && y.name.common == null) return 0;
            else if (x.name.common == null) return -1;
            else if (y.name.common == null) return 1;
            else return x.name.common.CompareTo(y.name.common);
        });

	// IDENTIFY CURRENCY CODES
		for(int i=0; i<listOfCountries.Count; i++) {
			List<string> tempList = new List<string>();
			foreach( KeyValuePair<string, Currencies> kvp in listOfCountries[i].currencies){
				tempList.Add(kvp.Key);
			}
			listOfCountries[i].currencyCodes = tempList;
		}
	}

	// CREATE EXCEL SHEET
	static async Task MakeExcel() {
		var excelApp = new Excel.Application();
		excelApp.Visible = true;

		excelApp.Workbooks.Add();
		Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

	// CREATE TITLE
		workSheet.Range[workSheet.Cells[1,1], workSheet.Cells[1,4]].Merge();

		workSheet.Cells[1, "A"] = "Countries List";

	// CREATE HEADER
		workSheet.Cells[2,"A"] = "Name";
		workSheet.Cells[2,"B"] = "Capital";
		workSheet.Cells[2,"C"] = "Area";
		workSheet.Cells[2,"D"] = "Currencies";
		
	// SET TITLE AND HEADER STYLE
		workSheet.Cells[1, "A"].Font.Size = 16;
		workSheet.Cells[1, "A"].Font.Color = 0x4F4F4F;
		workSheet.Range[workSheet.Cells[1, "A"],workSheet.Cells[2,"D"]].Font.Bold = true;
		workSheet.Range[workSheet.Cells[1, "A"],workSheet.Cells[2,"D"]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
		workSheet.Range[workSheet.Cells[2, "A"],workSheet.Cells[2,"D"]].Font.Size = 12;
		workSheet.Range[workSheet.Cells[2, "A"],workSheet.Cells[2,"D"]].Font.Color = 0x808080;

	// FILL SHEET
		await FillExcel(workSheet);
	
	// FIT WORDS IN COLUMNS
		((Excel.Range)workSheet.Columns[1]).AutoFit();
		((Excel.Range)workSheet.Columns[2]).AutoFit();
		((Excel.Range)workSheet.Columns[3]).AutoFit();
		((Excel.Range)workSheet.Columns[4]).AutoFit();
	}

	// FILL EXCEL SHEET
	static async Task FillExcel(Excel._Worksheet ws) {
		int row = 3;
		foreach(Countries pais in listOfCountries) {
			string nameToWrite, capitalToWrite, currenciesToWrite;
			double? areaToWrite;
			
		// DEFINE LINES TO BE WRITTEN
			nameToWrite = pais.name.common;
			areaToWrite = pais.area;
			
			if(pais.capital.Count == 0) {
				capitalToWrite = "-";
			} 
			else {
				capitalToWrite = "";

				for(int j=0; j<pais.capital.Count;j++) {
					capitalToWrite += pais.capital[j];
					if(j+1 < pais.capital.Count) {
						capitalToWrite += ", ";
					}
					else {
						capitalToWrite += " ";
					}
				}
			}

			currenciesToWrite = "";
			if(pais.currencyCodes.Count == 0) {
				currenciesToWrite = "-";
			} 
			else {
				for(int j=0; j<pais.currencyCodes.Count;j++) {
					currenciesToWrite += pais.currencyCodes[j];
					if(j+1 < pais.currencyCodes.Count) {
						currenciesToWrite += ", ";
					}
					else {
						currenciesToWrite += " ";
					}
				}
			}

		// WRITE INFORMATION LINES
			ws.Cells[row,"A"] = nameToWrite;
			ws.Cells[row,"B"] = capitalToWrite;
			if(areaToWrite == null) {
				ws.Cells[row,"C"] = "-";
			}
			else{
				ws.Cells[row,"C"] = areaToWrite;
				ws.Cells[row,"C"].NumberFormat = "0.00";
			}
			ws.Cells[row,"D"] = currenciesToWrite;

			row++;
		}
	}
}