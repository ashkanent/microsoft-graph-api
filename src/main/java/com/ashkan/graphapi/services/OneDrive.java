package com.ashkan.graphapi.services;

import com.ashkan.graphapi.auth.GraphUser;
import com.microsoft.graph.models.extensions.DriveItem;
import com.microsoft.graph.models.extensions.IGraphServiceClient;
import com.microsoft.graph.requests.extensions.IDriveItemCollectionPage;

public class OneDrive {

	public static void listAllFiles(String accessToken) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);

		IDriveItemCollectionPage drivePage = graphClient.me().drive().root().children().buildRequest().get();
		for (DriveItem item : drivePage.getCurrentPage()) {
			System.out.println(item.name + " (" + item.id + ")");
		}
	}

	public static DriveItem getFileById(String accessToken, String fileId) {
		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);

		return graphClient.me().drive().items(fileId).buildRequest().get();
	}

//	public static void createExcelFile(String accessToken, String fileName) {
//		IGraphServiceClient graphClient = GraphUser.ensureGraphClient(accessToken);
//	}

}
