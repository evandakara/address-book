import javax.swing.*;

public class Runner {

    public static void main(String[] args) {
        System.setProperty("swing.defaultl", "com.sun.java.swing.plaf.nimbus.NimbusLookAndFeel");
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new Form();
            }
        });
    }
}
// STOPSHIP: 12/01/2020