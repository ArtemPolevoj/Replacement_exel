import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;

public class Files {

    static File[] open() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setAcceptAllFileFilterUsed(false);
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Exsel.xlsx", "xlsx"));
        int s = fileChooser.showOpenDialog(null);

        return fileChooser.getSelectedFiles();
    }

    public static String save() {

        JFileChooser saveFileChooser = new JFileChooser();

        saveFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        saveFileChooser.setDialogTitle("СОЗДАЙТЕ ИЛИ ВЫБЕРИТЕ ФАЙЛ ДЛЯ СОХРАНЕНИЯ");
        saveFileChooser.setMultiSelectionEnabled(true);
        saveFileChooser.setAcceptAllFileFilterUsed(false);
        saveFileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Exsel.xlsx", "xlsx"));


        int i =saveFileChooser.showSaveDialog(null);

        if(i==1)return "";

        else return saveFileChooser.getSelectedFile().getPath();


    }

    }

