import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;

public class SecParcer {

    public static void main(String[] args) throws Exception {


    }

    private static String readUrl(String urlString) throws Exception {
        BufferedReader reader = null;
        try {
            URLConnection connection = new URL(urlString).openConnection();
            connection.setRequestProperty("User-Agent", "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Mobile Safari/537.36");
            connection.connect();
            reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
            StringBuffer buffer = new StringBuffer();
            int read;
            char[] chars = new char[1024];
            while ((read = reader.read(chars)) != -1) buffer.append(chars, 0, read);

            return buffer.toString();
        } catch (FileNotFoundException e){
            return "None";
        }
        finally {
            if (reader != null) reader.close();
        }
    }

    public static JsonNode getSecJsonNode(String cik, String parameter) throws Exception {
        String request = readUrl(getRequestSec(cik, parameter));
        if (!request.equals("None")){
        ObjectMapper objectMapper = new ObjectMapper();
        return objectMapper.readTree(request);
    } else return null;
    }

    public static long secValue(int year, JsonNode objYear) throws Exception {;

        long result = -1;
        if (objYear!=null){
        for (int i = 0; i < objYear.get("units").get("shares").size(); i++) {
            if(objYear.get("units").get("shares").get(i).get("fy").asInt()==year &&
                    objYear.get("units").get("shares").get(i).get("form").asText().equals("10-K"))
            {
                result = objYear.get("units").get("shares").get(i).get("val").asLong();
            }
        }
        }
        return result;
    }
    public static String getRequestSec(String cik, String parameter){
        return "https://data.sec.gov/api/xbrl/companyconcept/CIK"+ cik + "/us-gaap/"+ parameter +".json";
    }
}
