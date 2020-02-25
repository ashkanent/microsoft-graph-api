package com.ashkan.graphapi.controller;

import com.ashkan.graphapi.auth.Authentication;
import com.ashkan.graphapi.auth.GraphUser;
import com.ashkan.graphapi.services.Excel;
import com.ashkan.graphapi.services.OneDrive;
import com.microsoft.graph.models.extensions.DriveItem;
import com.microsoft.graph.models.extensions.User;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;

import javax.annotation.PostConstruct;

@Controller
public class controller {

	private Authentication authentication;

	@Autowired
	public controller(Authentication authentication) {
		this.authentication = authentication;
	}

	@PostConstruct
	public void init() {
		final String accessToken = authentication.getUserAccessToken();
		User user = GraphUser.getUser(accessToken);

		System.out.println("Welcome " + user.displayName);
//		OneDrive.listAllFiles(accessToken);
		final String testFileId = "01AXVMZA3U5XQPAI3WABALMGDW2V3HU23Z";
		final String testSheetId = "{00000000-0001-0000-0000-000000000000}";
		DriveItem testFile = OneDrive.getFileById(accessToken, testFileId);
//		Excel.createNewWorksheetForWorkbook(accessToken, testFile, "Ashkan");
		Excel.printWorksheetsForWorkbook(accessToken, testFile);
		String cellValue = Excel.getCellValueForWorkbookByRowAndColumn(accessToken, testFile, testSheetId, 0, 0);
		System.out.println("Cell value is: " + cellValue);

	}
}
