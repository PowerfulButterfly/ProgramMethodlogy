package cn.com.zsq.library;

import java.awt.Color;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import javax.swing.JComboBox;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.awt.event.ActionEvent;

public class test {

	private JFrame frame;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					test window = new test();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public test() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 678, 500);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JLabel lblNewLabel = new JLabel("选择时间：");
		lblNewLabel.setBounds(6, 6, 75, 16);
		frame.getContentPane().add(lblNewLabel);
		
		JComboBox comboBox = new JComboBox();
		comboBox.addItem("20190506");
		comboBox.addItem("20190507");
		comboBox.addItem("20190508");
		comboBox.addItem("20190509");
		comboBox.setBounds(93, 2, 155, 27);
		frame.getContentPane().add(comboBox);
		
		JComboBox comboBox_1 = new JComboBox();
		comboBox_1.addItem("8:00");
		comboBox_1.addItem("9:00");
		comboBox_1.addItem("10:00");
		comboBox_1.addItem("11:00");
		comboBox_1.setBounds(245, 2, 92, 27);
		frame.getContentPane().add(comboBox_1);
		
		JButton button = new JButton("卫生间（男）");
		button.setBounds(6, 100, 117, 104);
		frame.getContentPane().add(button);
		
		JButton button_1 = new JButton("卫生间（女）");
		button_1.setBounds(6, 235, 117, 104);
		frame.getContentPane().add(button_1);
		JLabel lbla = new JLabel("同学A占用");
		lbla.setBounds(177, 100, 94, 39);
		frame.getContentPane().add(lbla);
		
		JLabel label = new JLabel("同学A占用");
		label.setBounds(290, 100, 94, 39);
		frame.getContentPane().add(label);
		
		JLabel label_1 = new JLabel("同学A占用");
		label_1.setBounds(419, 100, 94, 39);
		frame.getContentPane().add(label_1);
		
		JLabel label_2 = new JLabel("同学A占用");
		label_2.setBounds(177, 262, 94, 39);
		frame.getContentPane().add(label_2);
		
		JLabel label_3 = new JLabel("同学A占用");
		label_3.setBounds(290, 262, 94, 39);
		frame.getContentPane().add(label_3);
		
		JLabel label_4 = new JLabel("同学A占用");
		label_4.setBounds(419, 262, 94, 39);
		frame.getContentPane().add(label_4);
		
		JButton button_10 = new JButton("301");
		button_10.setBounds(555, 110, 117, 29);
		frame.getContentPane().add(button_10);
		
		JButton button_11 = new JButton("302");
		button_11.setBounds(555, 138, 117, 29);
		frame.getContentPane().add(button_11);
		
		JButton button_12 = new JButton("303");
		button_12.setBounds(555, 166, 117, 29);
		frame.getContentPane().add(button_12);
		
		JButton button_13 = new JButton("304");
		button_13.setBounds(555, 209, 117, 29);
		frame.getContentPane().add(button_13);
		
		JButton button_8 = new JButton("305");
		button_8.setBounds(555, 235, 117, 29);
		frame.getContentPane().add(button_8);
		
		JButton button_9 = new JButton("306");
		button_9.setBounds(555, 262, 117, 29);
		frame.getContentPane().add(button_9);
		
		JComboBox comboBox_2 = new JComboBox();
		comboBox_2.addItem("一层 图书学习空间A");
		comboBox_2.addItem("一层 图书学习空间B");
		comboBox_2.addItem("二层 图书学习空间A");
		comboBox_2.addItem("二层 图书学习空间B");
		comboBox_2.addItem("三层 图书学习空间A");
		comboBox_2.addItem("三层 图书学习空间B");
		comboBox_2.addItem("四层 图书学习空间A");
		comboBox_2.addItem("四层 图书学习空间B");
		comboBox_2.setBounds(505, 370, 167, 27);
		frame.getContentPane().add(comboBox_2);
		
		JButton button_7 = new JButton("座位6");
		button_7.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Workbook book, book1, book2;
				try {
					book = Workbook.getWorkbook(new File("/Users/pro/Desktop/ZuoWei.xls"));
					Sheet sheet = book.getSheet(0);
					Cell cell1, cell2, cell3, cell4, cell5,cell6,cell7;
					String label1 = label_4.getText();
					if(label1.equals("暂时空闲")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						String time = "";//时间
						String date = "";//日期
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("6")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("暂时空闲")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								time = cell3.getContents();
								date = cell2.getContents();
								//5月6日 周一
								//5月7日 周二
								//5月8日 周三
								//5月9日 周四
								if(date.equals("20190506")) {
									date = "周一";
								}else if(date.equals("20190507")) {
									date = "周二";
								}else if(date.equals("20190508")) {
									date = "周三";
								}else {
									date = "周四";
								}
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						//查班级表
						//查课程表
						book1 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet1 = book1.getSheet(0);
						Cell cellA, cellB;
						String classes = "";
						int j = 1;
						while(true) {
							cellA = sheet1.getCell(0,j);
							cellB = sheet1.getCell(1,j);
							if(cellA.getContents().equals(stu_name)) {
								classes = cellB.getContents();
								break;
							}else if(cellA.getContents().equals("")) {
								break;
							}
							j++;
						}
						book2 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet2 = book2.getSheet(0);
						Cell c1, c2, c3, c4, c5, c6;
						int k = 1;
						String coarse1 = "";
						String coarse2 = "";
						//根据date time 查课表
						if(date.equals("周一")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周二")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周三")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else{
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}						
					}
					//占用
					else if(label1.equals("占用")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("6")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("占用")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								System.out.println(stu_name);
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						JOptionPane.showMessageDialog(null, stu_name+"同学已占用，请您预定其他座位", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
					else {
						JOptionPane.showMessageDialog(null, "该座位暂无预定，欢迎使用", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
				} catch (BiffException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}	
			
			}
		});
		button_7.setBounds(396, 217, 117, 47);
		frame.getContentPane().add(button_7);
		
		JButton button_6 = new JButton("座位5");
		button_6.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				Workbook book, book1, book2;
				try {
					book = Workbook.getWorkbook(new File("/Users/pro/Desktop/ZuoWei.xls"));
					Sheet sheet = book.getSheet(0);
					Cell cell1, cell2, cell3, cell4, cell5,cell6,cell7;
					String label1 = label_3.getText();
					if(label1.equals("暂时空闲")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						String time = "";//时间
						String date = "";//日期
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("5")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("暂时空闲")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								time = cell3.getContents();
								date = cell2.getContents();
								//5月6日 周一
								//5月7日 周二
								//5月8日 周三
								//5月9日 周四
								if(date.equals("20190506")) {
									date = "周一";
								}else if(date.equals("20190507")) {
									date = "周二";
								}else if(date.equals("20190508")) {
									date = "周三";
								}else {
									date = "周四";
								}
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						//查班级表
						//查课程表
						book1 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet1 = book1.getSheet(0);
						Cell cellA, cellB;
						String classes = "";
						int j = 1;
						while(true) {
							cellA = sheet1.getCell(0,j);
							cellB = sheet1.getCell(1,j);
							if(cellA.getContents().equals(stu_name)) {
								classes = cellB.getContents();
								break;
							}else if(cellA.getContents().equals("")) {
								break;
							}
							j++;
						}
						book2 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet2 = book2.getSheet(0);
						Cell c1, c2, c3, c4, c5, c6;
						int k = 1;
						String coarse1 = "";
						String coarse2 = "";
						//根据date time 查课表
						if(date.equals("周一")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周二")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周三")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else{
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}						
					}
					//占用
					else if(label1.equals("占用")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("5")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("占用")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								System.out.println(stu_name);
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						JOptionPane.showMessageDialog(null, stu_name+"同学已占用，请您预定其他座位", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
					else {
						JOptionPane.showMessageDialog(null, "该座位暂无预定，欢迎使用", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
				} catch (BiffException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}	
			}
		});
		button_6.setBounds(272, 217, 117, 47);
		frame.getContentPane().add(button_6);
		
		JButton button_5 = new JButton("座位4");
		button_5.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Workbook book, book1, book2;
				try {
					book = Workbook.getWorkbook(new File("/Users/pro/Desktop/ZuoWei.xls"));
					Sheet sheet = book.getSheet(0);
					Cell cell1, cell2, cell3, cell4, cell5,cell6,cell7;
					String label1 = label_2.getText();
					if(label1.equals("暂时空闲")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						String time = "";//时间
						String date = "";//日期
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("4")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("暂时空闲")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								time = cell3.getContents();
								date = cell2.getContents();
								//5月6日 周一
								//5月7日 周二
								//5月8日 周三
								//5月9日 周四
								if(date.equals("20190506")) {
									date = "周一";
								}else if(date.equals("20190507")) {
									date = "周二";
								}else if(date.equals("20190508")) {
									date = "周三";
								}else {
									date = "周四";
								}
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						//查班级表
						//查课程表
						book1 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet1 = book1.getSheet(0);
						Cell cellA, cellB;
						String classes = "";
						int j = 1;
						while(true) {
							cellA = sheet1.getCell(0,j);
							cellB = sheet1.getCell(1,j);
							if(cellA.getContents().equals(stu_name)) {
								classes = cellB.getContents();
								break;
							}else if(cellA.getContents().equals("")) {
								break;
							}
							j++;
						}
						book2 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet2 = book2.getSheet(0);
						Cell c1, c2, c3, c4, c5, c6;
						int k = 1;
						String coarse1 = "";
						String coarse2 = "";
						//根据date time 查课表
						if(date.equals("周一")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周二")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周三")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else{
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}						
					}
					//占用
					else if(label1.equals("占用")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("4")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("占用")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								System.out.println(stu_name);
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						JOptionPane.showMessageDialog(null, stu_name+"同学已占用，请您预定其他座位", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
					else {
						JOptionPane.showMessageDialog(null, "该座位暂无预定，欢迎使用", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
				} catch (BiffException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}	
			
			}
		});
		button_5.setBounds(154, 217, 117, 47);
		frame.getContentPane().add(button_5);
	
		JButton button_4 = new JButton("座位3");
		button_4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Workbook book, book1, book2;
				try {
					book = Workbook.getWorkbook(new File("/Users/pro/Desktop/ZuoWei.xls"));
					Sheet sheet = book.getSheet(0);
					Cell cell1, cell2, cell3, cell4, cell5,cell6,cell7;
					String label1 = label_1.getText();
					if(label1.equals("暂时空闲")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						String time = "";//时间
						String date = "";//日期
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("3")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("暂时空闲")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								time = cell3.getContents();
								date = cell2.getContents();
								//5月6日 周一
								//5月7日 周二
								//5月8日 周三
								//5月9日 周四
								if(date.equals("20190506")) {
									date = "周一";
								}else if(date.equals("20190507")) {
									date = "周二";
								}else if(date.equals("20190508")) {
									date = "周三";
								}else {
									date = "周四";
								}
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						//查班级表
						//查课程表
						book1 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet1 = book1.getSheet(0);
						Cell cellA, cellB;
						String classes = "";
						int j = 1;
						while(true) {
							cellA = sheet1.getCell(0,j);
							cellB = sheet1.getCell(1,j);
							if(cellA.getContents().equals(stu_name)) {
								classes = cellB.getContents();
								break;
							}else if(cellA.getContents().equals("")) {
								break;
							}
							j++;
						}
						book2 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet2 = book2.getSheet(0);
						Cell c1, c2, c3, c4, c5, c6;
						int k = 1;
						String coarse1 = "";
						String coarse2 = "";
						//根据date time 查课表
						if(date.equals("周一")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周二")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周三")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else{
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}						
					}
					//占用
					else if(label1.equals("占用")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("3")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("占用")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								System.out.println(stu_name);
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						JOptionPane.showMessageDialog(null, stu_name+"同学已占用，请您预定其他座位", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
					else {
						JOptionPane.showMessageDialog(null, "该座位暂无预定，欢迎使用", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
				} catch (BiffException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}	
			}
		});
		button_4.setBounds(396, 138, 117, 47);
		frame.getContentPane().add(button_4);
		
		
		
		JButton button_3 = new JButton("座位2");
		button_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Workbook book, book1, book2;
				try {
					book = Workbook.getWorkbook(new File("/Users/pro/Desktop/ZuoWei.xls"));
					Sheet sheet = book.getSheet(0);
					Cell cell1, cell2, cell3, cell4, cell5,cell6,cell7;
					String label1 = label.getText();
					if(label1.equals("暂时空闲")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						String time = "";//时间
						String date = "";//日期
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("2")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("暂时空闲")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								time = cell3.getContents();
								date = cell2.getContents();
								//5月6日 周一
								//5月7日 周二
								//5月8日 周三
								//5月9日 周四
								if(date.equals("20190506")) {
									date = "周一";
								}else if(date.equals("20190507")) {
									date = "周二";
								}else if(date.equals("20190508")) {
									date = "周三";
								}else {
									date = "周四";
								}
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						//查班级表
						//查课程表
						book1 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet1 = book1.getSheet(0);
						Cell cellA, cellB;
						String classes = "";
						int j = 1;
						while(true) {
							cellA = sheet1.getCell(0,j);
							cellB = sheet1.getCell(1,j);
							if(cellA.getContents().equals(stu_name)) {
								classes = cellB.getContents();
								break;
							}else if(cellA.getContents().equals("")) {
								break;
							}
							j++;
						}
						book2 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet2 = book2.getSheet(0);
						Cell c1, c2, c3, c4, c5, c6;
						int k = 1;
						String coarse1 = "";
						String coarse2 = "";
						//根据date time 查课表
						if(date.equals("周一")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周二")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周三")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else{
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}						
					}
					//占用
					else if(label1.equals("占用")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("2")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("占用")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								System.out.println(stu_name);
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						JOptionPane.showMessageDialog(null, stu_name+"同学已占用，请您预定其他座位", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
					else {
						JOptionPane.showMessageDialog(null, "该座位暂无预定，欢迎使用", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
				} catch (BiffException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}	
			}
		});
		button_3.setBounds(272, 138, 117, 47);
		frame.getContentPane().add(button_3);
		
		
		
		JButton button_2 = new JButton("座位1");
		button_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				//
				Workbook book, book1, book2;
				try {
					book = Workbook.getWorkbook(new File("/Users/pro/Desktop/ZuoWei.xls"));
					Sheet sheet = book.getSheet(0);
					Cell cell1, cell2, cell3, cell4, cell5,cell6,cell7;
					String label1 = lbla.getText();
					if(label1.equals("暂时空闲")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						String time = "";//时间
						String date = "";//日期
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("1")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("暂时空闲")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								time = cell3.getContents();
								date = cell2.getContents();
								//5月6日 周一
								//5月7日 周二
								//5月8日 周三
								//5月9日 周四
								if(date.equals("20190506")) {
									date = "周一";
								}else if(date.equals("20190507")) {
									date = "周二";
								}else if(date.equals("20190508")) {
									date = "周三";
								}else {
									date = "周四";
								}
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						//查班级表
						//查课程表
						book1 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet1 = book1.getSheet(0);
						Cell cellA, cellB;
						String classes = "";
						int j = 1;
						while(true) {
							cellA = sheet1.getCell(0,j);
							cellB = sheet1.getCell(1,j);
							if(cellA.getContents().equals(stu_name)) {
								classes = cellB.getContents();
								break;
							}else if(cellA.getContents().equals("")) {
								break;
							}
							j++;
						}
						book2 = Workbook.getWorkbook(new File("/Users/pro/Desktop/student_class.xls"));
						Sheet sheet2 = book2.getSheet(0);
						Cell c1, c2, c3, c4, c5, c6;
						int k = 1;
						String coarse1 = "";
						String coarse2 = "";
						//根据date time 查课表
						if(date.equals("周一")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(1,1).getContents();//第一节课
								coarse2 = sheet2.getCell(1,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周二")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(2,1).getContents();//第一节课
								coarse2 = sheet2.getCell(2,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else if(date.equals("周三")) {
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(3,1).getContents();//第一节课
								coarse2 = sheet2.getCell(3,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}else{
							if(time.equals("8:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位4小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("9:00")) {
								coarse1 = sheet2.getCell(4,1).getContents();//第一节课
								coarse2 = sheet2.getCell(4,2).getContents();//第二节课
								//第一节有课，第二节没课
								if(!(coarse1.equals("")) && coarse2.equals("")) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
								//第一第二节都有课
								else if(!(coarse1.equals("")) && !(coarse2.equals(""))) {
									JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位3小时", "通知", JOptionPane.INFORMATION_MESSAGE);
								}
							}else if(time.equals("10:00")) {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位2小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}else {
								JOptionPane.showMessageDialog(null, stu_name+"同学已预订，"+"您还可以使用该座位1小时", "通知", JOptionPane.INFORMATION_MESSAGE);
							}
						}						
					}
					//占用
					else if(label1.equals("占用")) {
						//记录 座位号1
						//记录 日期 时间 学习空间
						String comboBoxStr = comboBox.getSelectedItem().toString();//日期
						String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
						String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
						int i = 1;
						String stu_name = "";//学生姓名
						while(true) {
							cell1 = sheet.getCell(0,i);//座位
							cell2 = sheet.getCell(1,i);//日期
							cell3 = sheet.getCell(2,i);//时间
							cell4 = sheet.getCell(3,i);//状态
							cell5 = sheet.getCell(4,i);//学习空间
							cell6 = sheet.getCell(5,i);//教室
							cell7 = sheet.getCell(6,i);//学生姓名
							if((cell1.getContents().equals("1")) && (cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell4.getContents().equals("占用")) && (cell5.getContents().equals(comboBoxStr2))) {
								stu_name = cell7.getContents();
								System.out.println(stu_name);
								break;
							}else if("".equals(cell1.getContents())){
								break;
							}
							i++;
						}
						JOptionPane.showMessageDialog(null, stu_name+"同学已占用，请您预定其他座位", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
					else {
						JOptionPane.showMessageDialog(null, "该座位暂无预定，欢迎使用", "通知", JOptionPane.INFORMATION_MESSAGE);
					}
				} catch (BiffException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				
				
				
				
				
				
			}
		});
		button_2.setBounds(154, 138, 117, 47);
		frame.getContentPane().add(button_2);
		
		
		JButton btnNewButton = new JButton("查询");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					Workbook book = Workbook.getWorkbook(new File("/Users/pro/Desktop/ZuoWei.xls"));
					Sheet sheet = book.getSheet(0);
					Cell cell1, cell2, cell3, cell4, cell5,cell6;
					String comboBoxStr = comboBox.getSelectedItem().toString();//日期
					String comboBoxStr1 = comboBox_1.getSelectedItem().toString();//时间
					String comboBoxStr2 = comboBox_2.getSelectedItem().toString();//学习空间
					int i = 1;
					int a = 0;
					String[] str = new String[100];
					while(true) {
						cell1 = sheet.getCell(0,i);//座位
						cell2 = sheet.getCell(1,i);//日期
						cell3 = sheet.getCell(2,i);//时间
						cell4 = sheet.getCell(3,i);//状态
						cell5 = sheet.getCell(4,i);//学习空间
						cell6 = sheet.getCell(5,i);//教室
						System.out.println(cell2.getContents()+" "+cell3.getContents()+" "+cell5.getContents());
						if((cell2.getContents().equals(comboBoxStr)) && (cell3.getContents().equals(comboBoxStr1)) && (cell5.getContents().equals(comboBoxStr2))) {
							str[a] = cell4.getContents();
							a++;
						}else if("".equals(cell1.getContents())){
							break;
						}
						i++;
					}
					
					if(str[0].equals("空闲")) {
						lbla.setText(str[0]);
						lbla.setForeground(Color.BLUE);
					}else if(str[0].equals("暂时空闲")) {
						lbla.setText(str[0]);
						lbla.setForeground(Color.ORANGE);
					}else if(str[0].equals("占用")) {
						lbla.setText(str[0]);
						lbla.setForeground(Color.RED);
					}
					
					if(str[1].equals("空闲")) {
						label.setText(str[1]);
						label.setForeground(Color.BLUE);
					}else if(str[1].equals("暂时空闲")) {
						label.setText(str[1]);
						label.setForeground(Color.ORANGE);
					}else if(str[1].equals("占用")) {
						label.setText(str[1]);
						label.setForeground(Color.RED);
					}
					
					if(str[2].equals("空闲")) {
						label_1.setText(str[2]);
						label_1.setForeground(Color.BLUE);
					}else if(str[2].equals("暂时空闲")) {
						label_1.setText(str[2]);
						label_1.setForeground(Color.ORANGE);
					}else if(str[2].equals("占用")) {
						label_1.setText(str[2]);
						label_1.setForeground(Color.RED);
					}
					
					if(str[3].equals("空闲")) {
						label_2.setText(str[3]);
						label_2.setForeground(Color.BLUE);
					}else if(str[3].equals("暂时空闲")) {
						label_2.setText(str[3]);
						label_2.setForeground(Color.ORANGE);
					}else if(str[3].equals("占用")) {
						label_2.setText(str[3]);
						label_2.setForeground(Color.RED);
					}
					
					if(str[4].equals("空闲")) {
						label_3.setText(str[4]);
						label_3.setForeground(Color.BLUE);
					}else if(str[4].equals("暂时空闲")) {
						label_3.setText(str[4]);
						label_3.setForeground(Color.ORANGE);
					}else if(str[4].equals("占用")) {
						label_3.setText(str[4]);
						label_3.setForeground(Color.RED);
					}
			
					if(str[5].equals("空闲")) {
						label_4.setText(str[5]);
						label_4.setForeground(Color.BLUE);
					}else if(str[5].equals("暂时空闲")) {
						label_4.setText(str[5]);
						label_4.setForeground(Color.ORANGE);
					}else if(str[5].equals("占用")) {
						label_4.setText(str[5]);
						label_4.setForeground(Color.RED);
					}
				} catch (BiffException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				
			}
		});
		btnNewButton.setBounds(349, 1, 75, 29);
		frame.getContentPane().add(btnNewButton);
		
	}
}
