package com.muscleape.excel.util;

import java.io.File;

/**
 * @author Muscleape
 */
public class POITempFile {

    private static final String JAVA_IO_TMPDIR = "java.io.tmpdir";

    private static final String POIFILES = "poifiles";

    /**
     * 在系统临时目录下新建POIFILES文件夹
     *
     * @return
     * @author Muscleape
     * @date 2019/4/28 13:52
     */
    public static void createPOIFilesDirectory() {

        String tmpDir = System.getProperty(JAVA_IO_TMPDIR);
        if (tmpDir == null) {
            throw new RuntimeException(
                    "Systems temporary directory not defined - set the -D" + JAVA_IO_TMPDIR + " jvm property!");
        }
        File directory = new File(tmpDir, POIFILES);
        if (!directory.exists()) {
            syncCreatePOIFilesDirectory(directory);
        }

    }

    /**
     * @param directory
     */
    private static synchronized void syncCreatePOIFilesDirectory(File directory) {
        if (!directory.exists()) {
            directory.mkdirs();
        }
    }
}
