
package com.sample;
 
import javax.jws.WebMethod;
import javax.jws.WebService;
 
@WebService
public interface SimpleMath {
 @WebMethod
 public long sum(long a, long b);
}