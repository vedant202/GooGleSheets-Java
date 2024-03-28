package com.GoogleSheets;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.SpreadsheetProperties;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

public class SheetSrvices {
	private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
	  private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
	  private static final String TOKENS_DIRECTORY_PATH = "tokens";

	 
	  private static final List<String> SCOPES =
      Collections.singletonList(SheetsScopes.SPREADSHEETS);
	  private static final String CREDENTIALS_FILE_PATH = "/tutorial-sheet-credentials.json";
	  
	  final String spreadsheetId = "1seQ2ABEQMIoKDsTh-PMBvbnd112mcQFnKBFrgXY4Tdw";
	    final String range = "Sheet1!A:F";
	  
	public static Sheets sheetService() throws GeneralSecurityException, IOException {
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
	    InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);

	    if (in == null) {
	      throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
	    }
	    GoogleCredential  googleCredential = GoogleCredential.fromStream(in).createScoped(SCOPES);
	    
	    Sheets service =
	        new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, googleCredential)
	            .setApplicationName(APPLICATION_NAME)
	            .build();
	    return service;
	}
	
	public void getData(Sheets service) throws IOException {
	    


		ValueRange valueRange = service.spreadsheets().values().get(spreadsheetId, range).execute();
	    
	    if(valueRange.getValues()==null) {
	    	System.out.println("No data found");
	    }
	    else {
	    	System.out.println(valueRange.getRange());
	    	valueRange.getValues().forEach((i)->{System.out.println(i);});
	    }
	}
	
	public UpdateValuesResponse writeData(Sheets service) throws IOException {
		String range2 = "Sheet1!A13:F13";
		ValueRange vr = new ValueRange();
		vr.setValues(List.of(List.of("test2","test2","test2","test2","test2","test2")));
		AppendValuesResponse appendValuesResponse = service.spreadsheets().values().append(spreadsheetId, range2, vr).setValueInputOption("RAW").execute();
		UpdateValuesResponse updateValuesResponse = appendValuesResponse.getUpdates();
		System.out.println("updateValuesResponse "+updateValuesResponse);
		return updateValuesResponse;
	}
	
	public String createSpreadSheets(Sheets service,String title) throws IOException {
		Spreadsheet spreadSheet = new Spreadsheet()
				.setProperties(new SpreadsheetProperties().setTitle(title));
		
		spreadSheet =service.spreadsheets().create(spreadSheet).setFields("spreadsheetId").execute();
		
		System.out.println("SpreadSheet Id :- "+spreadSheet.getSpreadsheetId());
		
		return spreadSheet.getSpreadsheetId();
	}
}
