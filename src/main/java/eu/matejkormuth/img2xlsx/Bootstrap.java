package eu.matejkormuth.img2xlsx;

import java.io.File;
import java.io.IOException;

public class Bootstrap {

    public static void main(String[] args) {
        String img = args[0];
        try {
            new Application(new File(img), new File(img + ".xlsx"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
