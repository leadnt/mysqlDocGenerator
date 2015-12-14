package com.guitar.dbdoc;

import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author hxy
 */
public class Main {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)  {
        try {
            File file = new File("config.cfg");
            if(!file.exists()){
                System.out.println("file not exists! " + file.getAbsolutePath() );
                //System.exit(0);
                CreateProperties.create(file);
            }
            DocFactory docfac = new DocFactory(file);
            docfac.generator();
        } catch (Exception ex) {
            ex.printStackTrace();
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
        System.out.println("generator success!");
    }
    
}
