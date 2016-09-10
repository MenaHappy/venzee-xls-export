package javaapplication3;

import java.io.IOException;
import org.json.simple.parser.ParseException;

public class JavaApplication3 {

    public static void main(String[] args) throws IOException, ParseException {
        venzee_export ve = new venzee_export();
        ve.export(args[0], args[1]);
        System.exit(1);
    }
    
}
