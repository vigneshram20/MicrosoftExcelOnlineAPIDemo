package MicrosoftOnlineExcelAPIDemo.MicrosoftOnlineExcelAPIDemo;


import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ClientCredentialParameters;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.models.User;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.requests.GraphServiceClient;
import com.nimbusds.oauth2.sdk.http.HTTPResponse;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Collections;
import java.util.Properties;


class MicrosoftExcelOnlineDemo {

    private static String authority;
    private static String clientId;
    private static String secret;
    private static String scope;
    private static String existingSpreadSheetID = "<Your Excel Sheet ID>";
    private static String tenantID = "<Your Tenant ID>";

    public static void main(String args[]) throws Exception{

        setUpSampleData();

        try {
            
        	getExcelSheetJSONObject(clientId, secret, tenantID);


        } catch(Exception ex){
            System.out.println("Oops! We have an exception of type - " + ex.getClass());
            System.out.println("Exception message - " + ex.getMessage());
            throw ex;
        }
    }

    private static void getExcelSheetJSONObject(String clientID, String secret, String tenantID) throws Exception {

    	

        // With client credentials flows the scope is ALWAYS of the shape "resource/.default", as the
        // application permissions need to be set statically (in the portal), and then granted by a tenant administrator

        
    	//Client Secret builder
        final ClientSecretCredential clientSecretCredential = new ClientSecretCredentialBuilder()
                .clientId(clientId)
                .clientSecret(secret)
                .tenantId("tenantID")
                .build();
        
        
        //Token Credential Auth Provider
        final TokenCredentialAuthProvider tokenCredentialAuthProvider = new TokenCredentialAuthProvider(clientSecretCredential);
        
        
        //Graph Service Client Builder
        GraphServiceClient graphClient = GraphServiceClient.builder().authenticationProvider(tokenCredentialAuthProvider).buildClient();
        
        //User user = graphClient.me().buildRequest().get();

		//Getting WorkbookRange from the Excel Sheet
		  WorkbookRange workbookRange =
		  graphClient.me().drive().items(existingSpreadSheetID).workbook().
		  worksheets("Sheet1") .usedRange() .buildRequest().get();
		  
		  //Converting the JSONResponse values into string
		  String response =workbookRange.values.toString();
		  
		  JsonElement jsonElement = new JsonParser().parse(response);
	      JsonObject jsonObject = jsonElement.getAsJsonObject();
	      
	      System.out.println( jsonObject.get("text") );
	        
        
    }


    /**
     * Helper function unique to this sample setting. In a real application these wouldn't be so hardcoded, for example
     * different users may need different authority endpoints or scopes
     */
    private static void setUpSampleData() throws IOException {
        // Load properties file and set properties used throughout the sample
        Properties properties = new Properties();
        properties.load(Thread.currentThread().getContextClassLoader().getResourceAsStream("application.properties"));
        authority = properties.getProperty("AUTHORITY");
        clientId = properties.getProperty("CLIENT_ID");
        secret = properties.getProperty("SECRET");
        scope = properties.getProperty("SCOPE");
    }
}
