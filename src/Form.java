import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

public class Form extends JFrame {

    // Form Attributes
    private JTextField id;
    private JComboBox region;
    private JTextField name;
    private JTextArea address;
    private JTextField housePhone;
    private JTextField cellphone1;
    private JTextField cellphone2;
    private JTextField email;
    private JTextField website;
    private JTextField socialMedia;

    // Form Labels
    private JLabel labelId;
    private JLabel labelRegion;
    private JLabel labelName;
    private JLabel labelAddress;
    private JLabel labelHousePhone;
    private JLabel labelCellphone1;
    private JLabel labelCellphone2;
    private JLabel labelEmail;
    private JLabel labelWebsite;
    private JLabel labelSocialMedia;

    // Form Buttons
    private JButton buttonNew;
    private JButton buttonSave;
    private JButton buttonDelete;
    private JButton buttonRefresh;
    private JButton buttonClear;

    // Miscellaneous
    private JTable table;
    private JMenuBar menuBar;
    private JScrollPane scrollTable;

    // Retrieve The Excel File
    private File excelFile = new File(Configuration.DATABASE_PATH);

    // Constructors
    public Form() {
        initComponent();
        loadData();
    }

    // Initialize Components
    private void initComponent() {
        generateMenu();
        generateButton();
        generateLabels();
        generateAttributes();
        generateTable();

        // Create Pane With a Null Layout
        JPanel contentPane = generateContentPane();

        // Add Custom Panels to Panel Root
        getContentPane().add(contentPane);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setTitle(Configuration.PROJECT_TITLE);
        setSize(1000, 780);
        // setLocationRelativeTo(null);
        setVisible(true);
        pack();
    }

    // Read / Get Data From Microsoft Excel and Set To Table
    private void loadData() {

    }

    // A Simple Method to Generate The Menu Bar
    private void generateMenu() {
        menuBar = new JMenuBar();

        JMenu menuFile = new JMenu("File");
        menuFile.setMnemonic('f');
        JMenuItem about = new JMenuItem("About");
        about.setMnemonic('a');
        JMenuItem exit = new JMenuItem("Exit");
        exit.setMnemonic('x');

        menuFile.add(about);
        menuFile.add(exit);

        menuBar.add(menuFile);

        setJMenuBar(menuBar);
    }

    // A Method to Generate and Define Button Behaviour
    private void generateButton() {

        buttonNew = new JButton();
        buttonNew.setBounds(30, 300, 90, 35);
        buttonNew.setBackground(new Color(255, 255, 255));
        buttonNew.setForeground(new Color(0, 0, 0));
        buttonNew.setEnabled(true);
        buttonNew.setFont(new Font("sansserif", 0, 12));
        buttonNew.setText("New");
        buttonNew.setVisible(true);
        buttonNew.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // TODO : CREATE ACTION LISTENER FOR NEW BUTTON
            }
        });

        buttonSave = new JButton();
        buttonSave.setBounds(130, 300, 90, 35);
        buttonSave.setBackground(new Color(255, 255, 255));
        buttonSave.setForeground(new Color(0, 0, 0));
        buttonSave.setEnabled(true);
        buttonSave.setFont(new Font("sansserif", 0, 12));
        buttonSave.setText("Save");
        buttonSave.setVisible(true);
        buttonSave.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // TODO : CREATE ACTION LISTENER FOR SAVE BUTTON
            }
        });

        buttonDelete = new JButton();
        buttonDelete.setBounds(230, 300, 90, 35);
        buttonDelete.setBackground(new Color(255, 255, 255));
        buttonDelete.setForeground(new Color(0, 0, 0));
        buttonDelete.setEnabled(true);
        buttonDelete.setFont(new Font("sansserif", 0, 12));
        buttonDelete.setText("Delete");
        buttonDelete.setVisible(true);
        buttonDelete.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // TODO : CREATE ACTION LISTENER FOR DELETE BUTTON
            }
        });

        buttonRefresh = new JButton();
        buttonRefresh.setBounds(330, 300, 90, 35);
        buttonRefresh.setBackground(new Color(255, 255, 255));
        buttonRefresh.setForeground(new Color(0, 0, 0));
        buttonRefresh.setEnabled(true);
        buttonRefresh.setFont(new Font("sansserif", 0, 12));
        buttonRefresh.setText("Refresh");
        buttonRefresh.setVisible(true);
        buttonRefresh.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // TODO : CREATE ACTION LISTENER FOR REFRESH BUTTON
            }
        });

        buttonClear = new JButton();
        buttonClear.setBounds(430, 300, 90, 35);
        buttonClear.setBackground(new Color(255, 255, 255));
        buttonClear.setForeground(new Color(0, 0, 0));
        buttonClear.setEnabled(true);
        buttonClear.setFont(new Font("sansserif", 0, 12));
        buttonClear.setText("Clear");
        buttonClear.setVisible(true);
        buttonClear.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // TODO : CREATE ACTION LISTENER FOR CLEAR BUTTON
            }
        });
    }

    // A Simple Method to Generate the Button's Labels
    private void generateLabels() {
        labelId = new JLabel();
        labelId.setBounds(30, 30, 90, 30);
        labelId.setBackground(new Color(255, 255, 255));
        labelId.setForeground(new Color(0, 0, 0));
        labelId.setEnabled(true);
        labelId.setFont(new Font("sansserif", 0, 12));
        labelId.setText("ID");
        labelId.setVisible(true);

        labelRegion = new JLabel();
        labelRegion.setBounds(30, 70, 90, 30);
        labelRegion.setBackground(new Color(255, 255, 255));
        labelRegion.setForeground(new Color(0, 0, 0));
        labelRegion.setEnabled(true);
        labelRegion.setFont(new Font("sansserif", 0, 12));
        labelRegion.setText("Region");
        labelRegion.setVisible(true);

        labelName = new JLabel();
        labelName.setBounds(30, 110, 90, 30);
        labelName.setBackground(new Color(255, 255, 255));
        labelName.setForeground(new Color(0, 0, 0));
        labelName.setEnabled(true);
        labelName.setFont(new Font("sansserif", 0, 12));
        labelName.setText("Name");
        labelName.setVisible(true);

        labelAddress = new JLabel();
        labelAddress.setBounds(30, 150, 90, 30);
        labelAddress.setBackground(new Color(255, 255, 255));
        labelAddress.setForeground(new Color(0, 0, 0));
        labelAddress.setEnabled(true);
        labelAddress.setFont(new Font("sansserif", 0, 12));
        labelAddress.setText("Address");
        labelAddress.setVisible(true);

        labelHousePhone = new JLabel();
        labelHousePhone.setBounds(500, 30, 90, 30);
        labelHousePhone.setBackground(new Color(255, 255, 255));
        labelHousePhone.setForeground(new Color(0, 0, 0));
        labelHousePhone.setEnabled(true);
        labelHousePhone.setFont(new Font("sansserif", 0, 12));
        labelHousePhone.setText("House Phone");
        labelHousePhone.setVisible(true);

        labelCellphone1 = new JLabel();
        labelCellphone1.setBounds(500, 70, 90, 30);
        labelCellphone1.setBackground(new Color(255, 255, 255));
        labelCellphone1.setForeground(new Color(0, 0, 0));
        labelCellphone1.setEnabled(true);
        labelCellphone1.setFont(new Font("sansserif", 0, 12));
        labelCellphone1.setText("Cellphone 1");
        labelCellphone1.setVisible(true);

        labelCellphone2 = new JLabel();
        labelCellphone2.setBounds(500, 110, 90, 30);
        labelCellphone2.setBackground(new Color(255, 255, 255));
        labelCellphone2.setForeground(new Color(0, 0, 0));
        labelCellphone2.setEnabled(true);
        labelCellphone2.setFont(new Font("sansserif", 0, 12));
        labelCellphone2.setText("Cellphone 2");
        labelCellphone2.setVisible(true);

        labelEmail = new JLabel();
        labelEmail.setBounds(500, 150, 90, 30);
        labelEmail.setBackground(new Color(255, 255, 255));
        labelEmail.setForeground(new Color(0, 0, 0));
        labelEmail.setEnabled(true);
        labelEmail.setFont(new Font("sansserif", 0, 12));
        labelEmail.setText("Email");
        labelEmail.setVisible(true);

        labelWebsite = new JLabel();
        labelWebsite.setBounds(500, 190, 90, 30);
        labelWebsite.setBackground(new Color(255, 255, 255));
        labelWebsite.setForeground(new Color(0, 0, 0));
        labelWebsite.setEnabled(true);
        labelWebsite.setFont(new Font("sansserif", 0, 12));
        labelWebsite.setText("Website");
        labelWebsite.setVisible(true);

        labelSocialMedia = new JLabel();
        labelSocialMedia.setBounds(500, 230, 90, 30);
        labelSocialMedia.setBackground(new Color(255, 255, 255));
        labelSocialMedia.setForeground(new Color(0, 0, 0));
        labelSocialMedia.setEnabled(true);
        labelSocialMedia.setFont(new Font("sansserif", 0, 12));
        labelSocialMedia.setText("Social Medias");
        labelSocialMedia.setVisible(true);
    }

    // A Simple Method to Generate the Form's Attributes
    private void generateAttributes() {

    }

    // A Simple Method to Generate the Table
    private void generateTable() {

    }

    // A Function to Generate Content Panes
    public JPanel generateContentPane() {
        return new JPanel();
    }
}