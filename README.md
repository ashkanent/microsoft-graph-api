#Microsoft Graph API

In this project I am trying to explore the *Microsoft Graph API* and more specifically its
endpoints for Excel and OneDrive. In its current form this repo is more like a work in progress
but it is built in a way to be expandable in future. This project is powered by Spring Boot and
meant to be used as a facade for other clients who want to work with Graph API.

Current focus is on exploring Graph API and from early investigations it seems like the provided
functionality is limited. I use the recommended [Java SDK](https://github.com/microsoftgraph/msgraph-sdk-java) to interact with this API. Other obstacle here
is this SDK and lack of documentation. If I can find workarounds for the discovered issues I will clean-up the 
services to be usable by other projects who wish to interact with the API. 