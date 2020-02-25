package com.ashkan.graphapi.auth;

import com.microsoft.graph.logger.DefaultLogger;
import com.microsoft.graph.logger.LoggerLevel;
import com.microsoft.graph.models.extensions.IGraphServiceClient;
import com.microsoft.graph.models.extensions.User;
import com.microsoft.graph.requests.extensions.GraphServiceClient;

public class GraphUser {
	private static IGraphServiceClient graphClient = null;
	private static SimpleAuthProvider simpleAuthProvider = null;

	public static IGraphServiceClient ensureGraphClient(String accessToken) {
		if (graphClient == null) {
			// Create the auth provider
			simpleAuthProvider = new SimpleAuthProvider(accessToken);

			// Create default logger to only log errors
			DefaultLogger logger = new DefaultLogger();
			logger.setLoggingLevel(LoggerLevel.ERROR);

			// Build a Graph client
			graphClient = GraphServiceClient.builder()
					.authenticationProvider(simpleAuthProvider)
					.logger(logger)
					.buildClient();

		}

		return graphClient;
	}

	public static User getUser(String accessToken) {
		ensureGraphClient(accessToken);

		// GET /me to get authenticated user
		User me = graphClient
				.me()
				.buildRequest()
				.get();

		return me;
	}
}
