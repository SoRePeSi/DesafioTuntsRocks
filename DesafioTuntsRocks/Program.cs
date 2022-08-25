using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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

		//await MakeExcel();

		string filePath = @"C:\resultFolder";
		Console.WriteLine(filePath);

		FileInfo newFile = new FileInfo(filePath + @"\countriesList.xlsx");
		if (newFile.Exists){
			newFile.Delete(); 
			newFile = new FileInfo(filePath + @"\countriesList.xlsx");
		}

		using (ExcelPackage package = new ExcelPackage(newFile)){
			ExcelWorksheet workSheet = package.Workbook.Worksheets.Add("Countries");

		// CREATE TITLE
			workSheet.Cells[1,1,1,4].Merge = true;

			workSheet.Cells[1, 1].Value = "Countries List";

		// CREATE HEADER
			workSheet.Cells[2,1].Value = "Name";
			workSheet.Cells[2,2].Value = "Capital";
			workSheet.Cells[2,3].Value = "Area";
			workSheet.Cells[2,4].Value = "Currencies";
		
		// SET TITLE AND HEADER STYLE
			workSheet.Cells[1, 1].Style.Font.Size = 16;
			workSheet.Cells[1, 1].Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#4F4F4F"));
			workSheet.Cells[1, 1, 2, 4].Style.Font.Bold = true;
			workSheet.Cells[1, 1, 2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			workSheet.Cells[2, 1, 2, 4].Style.Font.Size = 12;
			workSheet.Cells[2, 1, 2, 4].Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));

		// WRITE COUNTRIES' INFORMATION
			await FillSheet(workSheet);
			workSheet.Column(1).AutoFit();
			workSheet.Column(2).AutoFit();
			workSheet.Column(3).AutoFit();
			workSheet.Column(4).AutoFit();
			
		// SAVE
			package.Workbook.Properties.Title = "CountriesList";
			if (!Directory.Exists(@"C:\resultFolder")){
				Directory.CreateDirectory(@"C:\resultFolder");
			}
			package.Save();

		// OPEN FILE EXPLORER
			System.Diagnostics.Process.Start("explorer.exe", "/select, \"" + filePath + @"\countriesList.xlsx""");
		}
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

	// FILL SHEET
	static async Task FillSheet(ExcelWorksheet ws) {
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
			ws.Cells[row,1].Value = nameToWrite;
			ws.Cells[row,2].Value = capitalToWrite;
			if(areaToWrite == null) {
				ws.Cells[row,3].Value = "-";
			}
			else{
				ws.Cells[row,3].Value = areaToWrite;
				ws.Cells[row,3].Style.Numberformat.Format = "0.00";
			}
			ws.Cells[row,4].Value = currenciesToWrite;

			row++;
		}
	}
}