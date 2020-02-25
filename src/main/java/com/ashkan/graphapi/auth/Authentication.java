package com.ashkan.graphapi.auth;

import com.microsoft.aad.msal4j.DeviceCode;
import com.microsoft.aad.msal4j.DeviceCodeFlowParameters;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.microsoft.aad.msal4j.PublicClientApplication;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.net.MalformedURLException;
import java.util.Set;
import java.util.function.Consumer;

@Component
public class Authentication {

	// Set authority to allow only organizational accounts
	// Device code flow only supports organizational accounts
	private final String authority = "https://login.microsoftonline.com/common/";

	@Value("${app.id}")
	private String applicationId;

	@Value("#{new java.util.HashSet(T(java.util.Arrays).asList('${app.scopes}'))}")
	private Set<String> scopeSet;


	public String getUserAccessToken() {
		if (applicationId == null) {
			System.out.println("You must initialize Authentication before calling getUserAccessToken");
			return null;
		}

		PublicClientApplication app;
		try {
			// Build the MSAL application object with
			// app ID and authority
			app = PublicClientApplication.builder(applicationId)
					.authority(authority)
					.build();
		} catch (MalformedURLException e) {
			return null;
		}

		// Create consumer to receive the DeviceCode object
		// This method gets executed during the flow and provides
		// the URL the user logs into and the device code to enter
		Consumer<DeviceCode> deviceCodeConsumer = (DeviceCode deviceCode) -> {
			// Print the login information to the console
			System.out.println(deviceCode.message());
		};

		// Request a token, passing the requested permission scopes
		IAuthenticationResult result = app.acquireToken(
				DeviceCodeFlowParameters
						.builder(scopeSet, deviceCodeConsumer)
						.build()
		).exceptionally(ex -> {
			System.out.println("Unable to authenticate - " + ex.getMessage());
			return null;
		}).join();

		if (result != null) {
			return result.accessToken();
		}

		return null;
	}
}
