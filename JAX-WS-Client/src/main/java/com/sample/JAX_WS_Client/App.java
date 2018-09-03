package com.sample.JAX_WS_Client;

import java.net.MalformedURLException;
import java.net.URL;

import javax.xml.ws.BindingProvider;

import javax.xml.ws.WebServiceException;

import com.sample.client.MathWS;
import com.sample.client.SimpleMath;
//import com.sun.xml.internal.ws.client.BindingProviderProperties;


public class App 
{
    public static void main( String[] args )
    {
    	
    	URL url = null;
        WebServiceException e = null;
        try {
            url = new URL("http://localhost:8080/JAX-WS-EP-1?wsdl");
        } catch (MalformedURLException ex) {
            e = new WebServiceException(ex);
        }
    	MathWS mathWS = new MathWS(url);
    	    	 	
    	SimpleMath math = mathWS.getMathWSPort();
    	
    	//Set timeout until a connection is established
    	//((BindingProvider)math).getRequestContext().put(BindingProviderProperties.CONNECT_TIMEOUT, 10000);
    	((BindingProvider)math).getRequestContext().put("com.sun.xml.internal.ws.connect.timeout", 10000);
    	

    	//Set timeout until the response is received
    	//((BindingProvider) math).getRequestContext().put(BindingProviderProperties.REQUEST_TIMEOUT, 1000);
    	((BindingProvider)math).getRequestContext().put("com.sun.xml.internal.ws.request.timeout", 10000);
    	System.out.println("SUM : " + math.sum(5, 5));
    	
    }
}
