package com.exitsoft;

import java.io.File;

/**
 * Created by nayunhwan on 2016. 2. 21..
 */
public class MyFilter extends javax.swing.filechooser.FileFilter {


    String type;
    String desc;

    public MyFilter(String type, String desc){
        this.type = type;
        this.desc = desc;
    }

    @Override
    public boolean accept(File f) {
        return f.getName().endsWith(type) || f.isDirectory();
    }

    public String getType(){
        return type;
    }
    @Override
    public String getDescription() {
        return desc;
    }
}
