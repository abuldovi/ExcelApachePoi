import java.io.File;
import java.io.IOException;
import java.util.*;

import com.fasterxml.jackson.databind.ObjectMapper;


public class Parcer {
    private File file;




    public Parcer(File file) throws IOException {
            this.file = file;

    }

        public List parceFile() throws IOException {
            return new ObjectMapper().readValue(file, List.class);
        }
        public LinkedHashMap parce(List list, int i) throws IOException {


            return new ObjectMapper().convertValue(list.get(i), LinkedHashMap.class);
        }
}
