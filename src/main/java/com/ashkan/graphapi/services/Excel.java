package com.ashkan.graphapi.services;

import com.ashkan.graphapi.auth.GraphUser;
import com.google.gson.JsonArray;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.extensions.DriveItem;
import com.microsoft.graph.models.extensions.IGraphServiceClient;
import com.microsoft.graph.models.extensions.WorkbookRange;
import com.microsoft.graph.models.extensions.WorkbookSessionInfo;
import com.microsoft.graph.models.extensions.WorkbookWorksheet;
import com.microsoft.graph.requests.extensions.IWorkbookWorksheetCollectionPage;

import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

public class Excel {

	public static String createSession(String accessToken, String fileId) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);

		WorkbookSessionInfo sessionInfo = graphClient.me().drive().items(fileId).workbook().createSession(true).buildRequest().post();

		return sessionInfo.id;
	}

	public static void printWorksheetsForWorkbook(String accessToken, DriveItem givenFile) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);

		IWorkbookWorksheetCollectionPage worksheets = graphClient.me().drive().items(givenFile.id).workbook().worksheets().buildRequest().get();
		for (WorkbookWorksheet worksheet : worksheets.getCurrentPage()) {
			System.out.println(worksheet.name + " (" + worksheet.id + ")");
		}
	}

	public static void createNewWorksheetForWorkbook(String accessToken, DriveItem givenFile, String sheetName) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);

		WorkbookWorksheet newWorkbookWorksheet = new WorkbookWorksheet();
		newWorkbookWorksheet.name = sheetName;
		graphClient.me().drive().items(givenFile.id).workbook().worksheets().buildRequest().post(newWorkbookWorksheet);
	}

	public static WorkbookWorksheet getWorksheetFromWorkbookById(String accessToken, DriveItem givenFile, String sheetId) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);
		WorkbookWorksheet workbookWorksheet = new WorkbookWorksheet();
		try {
			workbookWorksheet =
					graphClient.me().drive().items(givenFile.id).workbook().worksheets(URLEncoder.encode(sheetId,
							StandardCharsets.UTF_8.toString())).buildRequest().get();
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		return workbookWorksheet;
	}

	public static String getCellValueForWorkbookByRowAndColumn(String accessToken, String fileId, String sheetId, Integer rowIndex, Integer columnIndex) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);
		WorkbookRange range = new WorkbookRange();

		try {
			range = graphClient.me().drive().items(fileId).workbook().worksheets(URLEncoder.encode(sheetId, StandardCharsets.UTF_8.toString())).cell(rowIndex, columnIndex).buildRequest().get();
		} catch (ClientException | UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		return ((JsonArray)((JsonArray)range.values).get(0)).get(0).toString();
	}

	public static void setCellValueForWorkbookByRowAndColumn(String accessToken, String fileId, String sheetId, String targetRange, WorkbookRange updatedRange) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);
		WorkbookRange range = new WorkbookRange();

		try {
			graphClient.me().drive().items(fileId).workbook().worksheets(URLEncoder.encode(sheetId, StandardCharsets.UTF_8.toString())).range(targetRange).buildRequest().patch(updatedRange);
		} catch (ClientException | UnsupportedEncodingException e) {
			e.printStackTrace();
		}
	}
}
