package com.sample;
 
import javax.jws.WebService;
 
@WebService(endpointInterface = "com.sample.SimpleMath", serviceName = "MathWS")
 
public class MathWS implements SimpleMath 
{     
 
  public long sum(long a, long b)
   {
    System.out.println("Summing "+a+" + "+b);
    return a+b;
   }
}