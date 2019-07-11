# How-To-Export-Jtable-Data-To-Excel-In-Java-Netbeans-GUI-Tutorial
How To Export Jtable Data To Excel In Java Netbeans GUI Tutorial
For More Info Visit - http://mauricemuteti.info/how-to-export-jtable-data-to-excel-in-java-netbeans-gui-tutorial/
<pre>

           FileOutputStream excelFOU = null;
        BufferedOutputStream excelBOU = null;
        XSSFWorkbook excelJTableExporter = null;

        //Choose Location For Saving Excel File
        JFileChooser excelFileChooser = new JFileChooser("C:\\Users\\Authentic\\Desktop");
//        Change Dilog Box Title
        excelFileChooser.setDialogTitle("Save As");
//        Onliny filter files with these extensions "xls", "xlsx", "xlsm"
        FileNameExtensionFilter fnef = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx", "xlsm");
        excelFileChooser.setFileFilter(fnef);
        int excelChooser = excelFileChooser.showSaveDialog(null);

//        Check if save button is clicked
        if (excelChooser == JFileChooser.APPROVE_OPTION) {
            
            try {
                //Import excel poi libraries to netbeans
                excelJTableExporter = new XSSFWorkbook();
                XSSFSheet excelSheet = excelJTableExporter.createSheet("JTable Sheet");
                //            Loop to get jtable columns and rows
                for (int i = 0; i < model.getRowCount(); i++) {
                    XSSFRow excelRow = excelSheet.createRow(i);
                    for (int j = 0; j < model.getColumnCount(); j++) {
                        XSSFCell excelCell = excelRow.createCell(j);

                        //Now Get ImageNames From JLabel
                        //get the last column
                        if (j == model.getColumnCount() - 1) {
                            JLabel excelJL = (JLabel) model.getValueAt(i, j);
                            ImageIcon excelIMageIcon = (ImageIcon) excelJL.getIcon();
                            //Before retrieving the description of the image first set it...
                            excelImagePath = excelIMageIcon.getDescription();
                        }
                        excelCell.setCellValue(model.getValueAt(i, j).toString());

//                        Change the values of the fourth column to image paths
                        if (excelCell.getColumnIndex() == model.getColumnCount() - 1) {
                            excelCell.setCellValue(excelImagePath);
                        }
                    }
                }   //Append xlsx file extensions to selected files. To create unique file names
                excelFOU = new FileOutputStream(excelFileChooser.getSelectedFile() + ".xlsx");
                excelBOU = new BufferedOutputStream(excelFOU);
                excelJTableExporter.write(excelBOU);
                JOptionPane.showMessageDialog(null, "Exported Successfully !!!........");
            } catch (FileNotFoundException ex) {
                ex.printStackTrace();
            } catch (IOException ex) {
                ex.printStackTrace();
            } finally {
                try {
                    if (excelBOU != null) {
                        excelBOU.close();
                    }
                    if (excelFOU != null) {
                        excelFOU.close();
                    }
                    if (excelJTableExporter != null) {
                        excelJTableExporter.close();
                    }
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
            
        }

</pre>
