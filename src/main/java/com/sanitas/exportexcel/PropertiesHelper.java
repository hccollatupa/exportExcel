package com.sanitas.exportexcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.apache.log4j.Logger;

public class PropertiesHelper {
	private static final Logger logger = Logger.getLogger(App.class);
	
    public static String getParameter(String key){
        String value = "";
        Properties prop = new Properties();
        String nombreArchivoPropiedades = "exportexcel.properties";
        InputStream input = null;
        try
        {
        	input = new FileInputStream(nombreArchivoPropiedades);
        	prop.load(input);
            value = prop.getProperty(key);
        }
        catch (IOException ex) {
        	ex.printStackTrace();
        	logger.error(ex);
        }
        finally
        {
        	if (input != null)
        	{
        		try
        		{
        			input.close();
                }
        		catch (IOException e)
        		{
        			e.printStackTrace();
                    logger.error(e);
                }
            }
        }
        return value;
    }
}