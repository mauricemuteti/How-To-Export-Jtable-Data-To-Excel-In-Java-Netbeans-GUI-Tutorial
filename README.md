# How-To-Export-Jtable-Data-To-Excel-In-Java-Netbeans-GUI-Tutorial
How To Export Jtable Data To Excel In Java Netbeans GUI Tutorial
For More Info Visit - http://mauricemuteti.info/how-to-export-jtable-data-to-excel-in-java-netbeans-gui-tutorial/
<pre>

    public void exportDataToExcel() {

        //First Download Apache POI Library For Dealing with excel files.
        //Then add the library to the current project
        FileOutputStream excelFos = null;
        XSSFWorkbook excelJTableExport = null;
        BufferedOutputStream excelBos = null;
        try {

            //Choosing Saving Location
            //Set default location to C:\Users\Authentic\Desktop or your preferred location
            JFileChooser excelFileChooser = new JFileChooser("C:\\Users\\Authentic\\Desktop");
            //Dialog box title
            excelFileChooser.setDialogTitle("Save As ..");
            //Filter only xls, xlsx, xlsm files
            FileNameExtensionFilter fnef = new FileNameExtensionFilter("Files", "xls", "xlsx", "xlsm");
            //Setting extension for selected file names
            excelFileChooser.setFileFilter(fnef);
            int chooser = excelFileChooser.showSaveDialog(null);
            //Check if save button has been clicked
            if (chooser == JFileChooser.APPROVE_OPTION) {
                //If button is clicked execute this code
                excelJTableExport = new XSSFWorkbook();
                XSSFSheet excelSheet = excelJTableExport.createSheet("Jtable Export");
                //Loop through the jtable columns and rows to get its values
                for (int i = 0; i < model.getRowCount(); i++) {
                    XSSFRow excelRow = excelSheet.createRow(i);
                    for (int j = 0; j < model.getColumnCount(); j++) {
                        XSSFCell excelCell = excelRow.createCell(j);

                        //Change the image column to output image path
                        //Fourth column contains images
                        if (j == model.getColumnCount() - 1) {
                            JLabel excelJL = (JLabel) model.getValueAt(i, j);
                            ImageIcon excelImageIcon = (ImageIcon) excelJL.getIcon();
                            //Image Name Is Stored In ImageIcons Description first set it when saving image in the jtable cell and then retrieve it.
                            excelImagePath = excelImageIcon.getDescription();
                        }

                        excelCell.setCellValue(model.getValueAt(i, j).toString());
                        if (excelCell.getColumnIndex() == model.getColumnCount() - 1) {
                            excelCell.setCellValue(excelImagePath);
                        }
                    }
                }
                excelFos = new FileOutputStream(excelFileChooser.getSelectedFile() + ".xlsx");
                excelBos = new BufferedOutputStream(excelFos);
                excelJTableExport.write(excelBos);
                JOptionPane.showMessageDialog(null, "Exported Successfully");
            }

        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(null, ex);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, ex);
        } finally {
            try {
                if (excelFos != null) {
                    excelFos.close();
                }
                if (excelBos != null) {
                    excelBos.close();
                }
                if (excelJTableExport != null) {
                    excelJTableExport.close();
                }
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, ex);
            }
        }
    } 
}

</pre>
