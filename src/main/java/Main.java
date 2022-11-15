import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class Main {

    public static void main(String[] args) throws IOException {


        JFrame frame = new JFrame("ÐÀÁÎÒÀ Ñ ÝÊÑÅËÜ");
        frame.setDefaultCloseOperation(3);
        frame.setResizable(false);
        frame.setLayout(new GridLayout(4, 2, 2, 10));
        frame.setLayout((LayoutManager) null);
        frame.setBounds(400, 120, 665, 520);

        JLabel labelReplace = new JLabel("Çàìåíèòü");
        labelReplace.setFont(new Font("Verdana", 1, 14));
        labelReplace.setBounds(10, 5, 110, 30);
        frame.add(labelReplace);

        JTextField fieldReplace = new JTextField();
        fieldReplace.setFont(new Font("Verdana", 1, 14));
        fieldReplace.setBounds(115, 5, 200, 30);
        frame.add(fieldReplace);

        JLabel labelReplacment = new JLabel("çàìåíèòü íà");
        labelReplacment.setFont(new Font("Verdana", 1, 14));
        labelReplacment.setBounds(330, 5, 110, 30);
        frame.add(labelReplacment);

        JTextField fieldReplacment = new JTextField();
        fieldReplacment.setFont(new Font("Verdana", 1, 14));
        fieldReplacment.setBounds(440, 5, 200, 30);
        frame.add(fieldReplacment);

        JTextArea areOutText = new JTextArea("Âûâîä ðåçóëüòàòîâ ðàáîòû:\n");
        areOutText.setFont(new Font("Verdana", 0, 14));
        areOutText.setBounds(10, 120, 630, 180);
        areOutText.setLineWrap(true);
        areOutText.setWrapStyleWord(true);
        frame.add(areOutText);

        JButton buttonReplace = new JButton("ÂÛÏÎËÍÈÒÜ ÇÀÌÅÍÓ");
        buttonReplace.setFont(new Font("Verdana", 1, 16));
        buttonReplace.setBounds(10, 50, 630, 50);
        buttonReplace.addActionListener(e -> {
                    ArrayList<File> openFile = Read.reading();
                    if (openFile.isEmpty()) {
                        areOutText.setText("Â âûáðàíîé ïàïêå íåò ýêñåëü ôàéëîâ.");
                    } else {
                        for (File file : openFile) {
                            areOutText.append(Replacement.replace(file, fieldReplace.getText(), fieldReplacment.getText()));
                        }
                    }
                });


        buttonReplace.setBorderPainted(true);
        frame.add(buttonReplace);
        JScrollPane scrollfieldOutText = new JScrollPane(areOutText);
        scrollfieldOutText.setVerticalScrollBarPolicy(20);
        scrollfieldOutText.setVisible(true);
        scrollfieldOutText.setBounds(10, 120, 630, 180);
        frame.add(scrollfieldOutText);

        JButton buttonMandatory = new JButton("ÏÎËÓ×ÈÒÜ ÎÇ");
        buttonMandatory.setFont(new Font("Verdana", Font.BOLD,20));
        buttonMandatory.setBorderPainted(true);
        buttonMandatory.setBounds(10,320,630,50);
        buttonMandatory.addActionListener(e -> areOutText.append(Mandatory.mandatory()));
        frame.add(buttonMandatory);

        JButton buttonExit = new JButton("ÂÛÕÎÄ");
        buttonExit.setFont(new Font("Verdana", 1, 20));
        buttonExit.setBorderPainted(true);
        buttonExit.setBounds(10, 420, 630, 50);
        buttonExit.addActionListener(e ->  System.exit(0));
        frame.add(buttonExit);
        frame.setVisible(true);
    }
}
