import com.sun.corba.se.spi.orbutil.threadpool.Work;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;

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
    private JScrollPane scrollAddress;

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
        if (excelFile.exists()) {
            try {
                Workbook workbook = Workbook.getWorkbook(excelFile);
                Sheet sheet = workbook.getSheets()[0];

                System.out.println("name " + sheet.getName());

                // To set Header Of Table
                String[] header = new String[]
                        {"ID", "REGION", "NAME", "ADDRESS", "HOUSEPHONE"
                                , "CELLPHONE1", "CELLPHONE2", "EMAIL", "WEBSITE", "SOCIAL MEDIA"};
            } catch (IOException | BiffException e) {
                e.printStackTrace();
            }
        } else {
            JOptionPane.showMessageDialog(null, String.valueOf("Error 404 : File not found" +
                    " contact @whxsbang for technical support/to report this issue"));
        }
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
                buttonNewActionPerformed();
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
                clearForm();
            }
        });
    }

    // A Method to Clear Form
    private void buttonNewActionPerformed() {
        newData();
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
        id = new JTextField();
        id.setBounds(100, 30, 100, 30);
        id.setBackground(new Color(255, 255, 255));
        id.setForeground(new Color(0, 0, 0));
        id.setEnabled(true);
        id.setEditable(false);
        id.setFont(new Font("sansserif", 0, 12));
        id.setText("");
        id.setVisible(true);

        region = new JComboBox();
        region.setBounds(100, 70, 150, 30);
        region.setBackground(new Color(255, 255, 255));
        region.setForeground(new Color(0, 0, 0));
        region.setEnabled(true);
        region.setEditable(false);
        region.setFont(new Font("sansserif", 0, 12));
        region.setModel(new DefaultComboBoxModel(new String[]{
                "Home", "School", "Online"
        }));
        region.setVisible(true);

        name = new JTextField();
        name.setBounds(100, 110, 220, 30);
        name.setBackground(new Color(255, 255, 255));
        name.setForeground(new Color(0, 0, 0));
        name.setEnabled(true);
        name.setEditable(true);
        name.setFont(new Font("sansserif", 0, 12));
        name.setText("");
        name.setVisible(true);

        address = new JTextArea();
        address.setBounds(100, 150, 220, 100);
        address.setBackground(new Color(255, 255, 255));
        address.setForeground(new Color(0, 0, 0));
        address.setEnabled(true);
        address.setEditable(true);
        address.setFont(new Font("sansserif", 0, 12));
        address.setText("");
        address.setVisible(true);
        address.setBorder(BorderFactory.createBevelBorder(1));
        address.setWrapStyleWord(true);
        address.setLineWrap(true);
        address.setAutoscrolls(true);

        scrollAddress = new JScrollPane();
        scrollAddress.setVisible(true);
        scrollAddress.setBounds(100, 150, 220, 100);
        scrollAddress.setViewportView(address);

        housePhone = new JTextField();
        housePhone.setBounds(620, 30, 220, 30);
        housePhone.setBackground(new Color(255, 255, 255));
        housePhone.setForeground(new Color(0, 0, 0));
        housePhone.setEnabled(true);
        housePhone.setEditable(true);
        housePhone.setFont(new Font("sansserif", 0, 12));
        housePhone.setText("");
        housePhone.setVisible(true);

        cellphone1 = new JTextField();
        cellphone1.setBounds(620, 70, 220, 30);
        cellphone1.setBackground(new Color(255, 255, 255));
        cellphone1.setForeground(new Color(0, 0, 0));
        cellphone1.setEnabled(true);
        cellphone1.setEditable(true);
        cellphone1.setFont(new Font("sansserif", 0, 12));
        cellphone1.setText("");
        cellphone1.setVisible(true);

        cellphone2 = new JTextField();
        cellphone2.setBounds(620, 110, 220, 30);
        cellphone2.setBackground(new Color(255, 255, 255));
        cellphone2.setForeground(new Color(0, 0, 0));
        cellphone2.setEnabled(true);
        cellphone2.setEditable(true);
        cellphone2.setFont(new Font("sansserif", 0, 12));
        cellphone2.setText("");
        cellphone2.setVisible(true);

        email = new JTextField();
        email.setBounds(620, 150, 220, 30);
        email.setBackground(new Color(255, 255, 255));
        email.setForeground(new Color(0, 0, 0));
        email.setEnabled(true);
        email.setEditable(true);
        email.setFont(new Font("sansserif", 0, 12));
        email.setText("");
        email.setVisible(true);

        website = new JTextField();
        website.setBounds(620, 190, 220, 30);
        website.setBackground(new Color(255, 255, 255));
        website.setForeground(new Color(0, 0, 0));
        website.setEnabled(true);
        website.setEditable(true);
        website.setFont(new Font("sansserif", 0, 12));
        website.setText("");
        website.setVisible(true);

        socialMedia = new JTextField();
        socialMedia.setBounds(620, 230, 220, 30);
        socialMedia.setBackground(new Color(255, 255, 255));
        socialMedia.setForeground(new Color(0, 0, 0));
        socialMedia.setEnabled(true);
        socialMedia.setEditable(true);
        socialMedia.setFont(new Font("sansserif", 0, 12));
        socialMedia.setText("");
        socialMedia.setVisible(true);
    }

    // A Simple Method to Generate the Table
    private void generateTable() {
        table = new JTable();
        table.setBounds(30, 370, 1300, 300);
        setLayout(new FlowLayout());
        table.setBackground(new Color(255, 255, 255));
        table.setFont(new Font("sansserif", 0, 12));
        table.setVisible(true);
        table.setRowHeight(25);
        table.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                tableMouseClicked(e);
            }
        });

        scrollAddress = new JScrollPane();
        scrollAddress.setVisible(true);
        scrollAddress.setBounds(100, 150, 220, 100);
        scrollAddress.setViewportView(address);
    }

    // A Function to Generate Content Panes
    public JPanel generateContentPane() {
        JPanel contentPane = new JPanel(null);
        contentPane.setPreferredSize(new Dimension(1360, 768));
        contentPane.setBackground(new Color(38, 35, 38, 242));

        // Adding buttons and labels into the proper panes
        contentPane.add(buttonClear);
        contentPane.add(buttonDelete);
        contentPane.add(buttonNew);
        contentPane.add(buttonRefresh);
        contentPane.add(buttonSave);
        contentPane.add(labelId);
        contentPane.add(labelRegion);
        contentPane.add(labelName);
        contentPane.add(labelAddress);
        contentPane.add(labelHousePhone);
        contentPane.add(labelCellphone1);
        contentPane.add(labelCellphone2);
        contentPane.add(labelEmail);
        contentPane.add(labelWebsite);
        contentPane.add(labelSocialMedia);

        contentPane.add(id);
        contentPane.add(region);
        contentPane.add(name);
        contentPane.add(address);
        contentPane.add(housePhone);
        contentPane.add(cellphone1);
        contentPane.add(cellphone2);
        contentPane.add(email);
        contentPane.add(website);
        contentPane.add(socialMedia);
        contentPane.add(scrollAddress);
        contentPane.add(table);

        return contentPane;
    }

    private void tableMouseClicked(MouseEvent event) {

    }

    private void newData() {
        clearForm();
        loadData();
    }

    private void clearForm() {
        id.setText(null);
        region.setSelectedIndex(0);
        name.setText(null);
        address.setText(null);
        housePhone.setText(null);
        cellphone1.setText(null);
        cellphone2.setText(null);
        email.setText(null);
        website.setText(null);
        socialMedia.setText(null);
    }
}