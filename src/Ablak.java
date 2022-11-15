import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import com.spire.data.table.DataTable;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import javax.swing.ButtonGroup;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;

import javax.swing.JTextField;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.JTextArea;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JButton;
import javax.swing.JFileChooser;

public class Ablak extends JFrame 
{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private JTextField ip;
	private JTextField user;
	private JTextField password;
	private JTextArea sql;
	private JFileChooser megnyit = new JFileChooser();
	private File excel;
	private JRadioButton radio_igen;
	private JRadioButton radio_nem;
	private JScrollPane gorgeto;
	private JScrollPane gorgeto2;
	private JTextArea letrehoz;
	private JRadioButton letrehoz_igen;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) 
	{
		EventQueue.invokeLater(new Runnable() 
		{
			public void run() 
			{
				try 
				{
					Ablak frame = new Ablak();
					frame.setVisible(true);
				} 
				catch (Exception e) 
				{
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public Ablak() 
	{
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 922, 578);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		
		setTitle("Adatbázis feltöltő");
		
		JLabel lblNewLabel = new JLabel("Adatb\u00E1zis felt\u00F6lt\u00E9se");
		lblNewLabel.setFont(new Font("Arial", Font.BOLD, 14));
		
		JLabel lblNewLabel_1 = new JLabel("IP:");
		
		ip = new JTextField();
		ip.setColumns(10);
		
		user = new JTextField();
		user.setColumns(10);
		
		JLabel lblNewLabel_2 = new JLabel("Felhaszn\u00E1l\u00F3:");
		
		password = new JTextField();
		password.setColumns(10);
		
		JLabel lblNewLabel_3 = new JLabel("Jelsz\u00F3:");
		
		JLabel lblNewLabel_4 = new JLabel("Fel SQL:");
		
		JLabel lblNewLabel_5 = new JLabel("Fejl\u00E9c:");
		
		radio_igen = new JRadioButton("Igen");
		radio_igen.setSelected(true);
		
		radio_nem = new JRadioButton("Nem");
		
		ButtonGroup csoport = new ButtonGroup();
		csoport.add(radio_igen);
		csoport.add(radio_nem);
		
		JButton megnyit = new JButton("Megnyit\u00E1s");
		megnyit.addActionListener(new Megnyitas());
		
		JButton feltolt = new JButton("Felt\u00F6lt");
		feltolt.addActionListener(new Feltolt());
		
		sql = new JTextArea();
		sql.setText("INSERT INTO adatbazis VALUE(?, ?, ?)");
		gorgeto = new JScrollPane(sql);
		
		letrehoz = new JTextArea();
		gorgeto2 = new JScrollPane(letrehoz);
		letrehoz.setText("CREATE TABLE pelda_tabla \n" +
                   "(id INTEGER not NULL, \n" +
                   " first VARCHAR(255),\n" + 
                   " age INTEGER, \n" + 
                   " PRIMARY KEY ( id ))");
		
		JLabel lblNewLabel_6 = new JLabel("Új tábla");
		
		JLabel lblNewLabel_7 = new JLabel("Létrehoz?");
		
		letrehoz_igen = new JRadioButton("Igen");
		
		JRadioButton letrehoz_nem = new JRadioButton("Nem");
		letrehoz_nem.setSelected(true);
		
		ButtonGroup csoport2 = new ButtonGroup();
		csoport2.add(letrehoz_igen);
		csoport2.add(letrehoz_nem);
		
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addGap(202)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addComponent(lblNewLabel_7)
							.addContainerGap())
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
										.addComponent(lblNewLabel_1)
										.addComponent(lblNewLabel_2)
										.addComponent(lblNewLabel_3)
										.addComponent(lblNewLabel_6)
										.addComponent(lblNewLabel_4)
										.addComponent(lblNewLabel_5))
									.addGap(30)
									.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
										.addComponent(radio_igen)
										.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
											.addComponent(ip, GroupLayout.PREFERRED_SIZE, 135, GroupLayout.PREFERRED_SIZE)
											.addGroup(gl_contentPane.createSequentialGroup()
												.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
													.addGroup(gl_contentPane.createSequentialGroup()
														.addComponent(gorgeto2, GroupLayout.DEFAULT_SIZE, 470, Short.MAX_VALUE)
														.addPreferredGap(ComponentPlacement.RELATED))
													.addComponent(gorgeto, GroupLayout.DEFAULT_SIZE, 470, Short.MAX_VALUE)
													.addGroup(gl_contentPane.createSequentialGroup()
														.addComponent(letrehoz_igen)
														.addPreferredGap(ComponentPlacement.RELATED))
													.addGroup(gl_contentPane.createSequentialGroup()
														.addComponent(letrehoz_nem)
														.addPreferredGap(ComponentPlacement.RELATED))
													.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING, false)
														.addComponent(password, Alignment.LEADING)
														.addComponent(user, Alignment.LEADING, GroupLayout.DEFAULT_SIZE, 236, Short.MAX_VALUE)))
												.addGap(171)))
										.addComponent(radio_nem)))
								.addComponent(lblNewLabel)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addComponent(megnyit)
									.addGap(78)
									.addComponent(feltolt)))
							.addGap(134))))
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addComponent(lblNewLabel)
					.addGap(38)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblNewLabel_1)
						.addComponent(ip, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(18)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblNewLabel_2)
						.addComponent(user, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(18)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblNewLabel_3)
						.addComponent(password, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGap(37)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(gorgeto2, GroupLayout.PREFERRED_SIZE, 56, GroupLayout.PREFERRED_SIZE)
						.addComponent(lblNewLabel_6))
					.addGap(26)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addComponent(lblNewLabel_7, GroupLayout.PREFERRED_SIZE, 29, GroupLayout.PREFERRED_SIZE)
						.addComponent(letrehoz_igen))
					.addGap(1)
					.addComponent(letrehoz_nem)
					.addGap(18)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(gorgeto, GroupLayout.PREFERRED_SIZE, 56, GroupLayout.PREFERRED_SIZE)
						.addComponent(lblNewLabel_4))
					.addGap(13)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addComponent(lblNewLabel_5, Alignment.TRAILING)
						.addComponent(radio_igen, Alignment.TRAILING))
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addComponent(radio_nem)
					.addGap(12)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(megnyit)
						.addComponent(feltolt))
					.addGap(20))
		);
		contentPane.setLayout(gl_contentPane);
	}
	
	class Megnyitas implements ActionListener																						//mentés gomb megnyomáskor hívodik meg
	{
		public void actionPerformed(ActionEvent e)
		 {
			megnyit.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);															//filechooser beállítása
		    megnyit.showOpenDialog(megnyit);																								//fc ablak megniytása
		    excel = megnyit.getSelectedFile();																					//kiválasztott fájl odaadása egy File osztálynak
		 }
	}
	
	class Feltolt implements ActionListener																						//mentés gomb megnyomáskor hívodik meg
	{
		public void actionPerformed(ActionEvent e)
		 {
			Connection conn = null;
	        Statement stmt = null;
	        try 
	        {
	           try 
	           {
	              Class.forName("com.mysql.cj.jdbc.Driver");                                //Driver meghívása
	           } 
	           catch (Exception e1) 
	           {
	              System.out.println(e1);
	              String hibauzenet2 = e1.toString();
	              JOptionPane.showMessageDialog(null, hibauzenet2, "Hiba üzenet", 2);
	           }
	           
	        conn = DriverManager.getConnection("jdbc:mysql://"+ ip.getText(), user.getText(), password.getText());                           //kapcsolat létrehozása
	        stmt = conn.createStatement();                                                                                                  //csatlakozás
	        
	        PreparedStatement statement = conn.prepareStatement(sql.getText());
	        PreparedStatement statement2 = conn.prepareStatement(letrehoz.getText());
	        
	        if(letrehoz_igen.isSelected())
	        {
	        	statement2.executeUpdate();
	        }
	        
	        Workbook workbook = new Workbook();
	        workbook.loadFromFile(excel.getAbsolutePath());
	        Worksheet sheet = workbook.getWorksheets().get(0);
	        
	        DataTable dataTable;
	        
	        if(radio_igen.isSelected())
	        { 
	        	dataTable = sheet.exportDataTable();
	        }
	        else
	        {
	        	dataTable = sheet.exportDataTable(sheet.getAllocatedRange(), false, false );                                  //sheet.getAllocatedRange(), false, false 	
	        }
	        
	        for (int i = 0; i < dataTable.getRows().size(); i++) 
	        {
	            for (int j = 0; j < dataTable.getColumns().size(); j++) 
	            {
	                statement.setString(j + 1, dataTable.getRows().get(i).getString(j));
	            }
	            statement.executeUpdate();
	        }
	        
	        statement.close();
	        conn.close();
	        
	        JOptionPane.showMessageDialog(null, "Feltöltés sikeres", "Info", 1);
	        } 
	        catch (SQLException e1)                                                     //kivétel esetén történik
	        {
	           e1.printStackTrace();
	           String hibauzenet2 = e1.toString();
	           JOptionPane.showMessageDialog(null, hibauzenet2 + "\n \n A Mentés sikertelen!!", "Hiba üzenet", 2);
	        } 
	        catch (Exception e1) 
	        {
	           e1.printStackTrace();
	           String hibauzenet2 = e1.toString();
	           JOptionPane.showMessageDialog(null, hibauzenet2, "Hiba üzenet", 2);
	        } 
	        finally                                                                     //finally rész mindenképpen lefut, hogy hiba esetén is lezárja a kacsolatot
	        {
	           try 
	           {
	              if (stmt != null)
	                 conn.close();
	           } 
	           catch (SQLException se) {}
	           try 
	           {
	              if (conn != null)
	                 conn.close();
	           } 
	           catch (SQLException se) 
	           {
	              se.printStackTrace();
	           }  
	        }	
		 }
	}
}
