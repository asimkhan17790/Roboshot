/*
 * 
 * Author: Asim Khan 
 * Date: 20th November 2013
 * 
 */
package com.acc;
import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.util.IOUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.awt.image.RenderedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

public class WindowApp extends JFrame implements ActionListener{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	/**
	 * @param args
	 */
	
	
	private List <byte[]> listImage = null;
	private JButton capture = null;
	private JButton createDoc = null;
	private JButton refresh = null;
	static JFrame frame = null;
	


	    
 public JPanel createContentPane(){    
	    	
	 JPanel totalGUI = new JPanel();
	 	totalGUI.setLayout(null);
	 	//Capture Button start
	 	//this.getClass().getClassLoader().getResourceAsStream("as");
	 	capture = new JButton(new ImageIcon( Toolkit.getDefaultToolkit().getImage(this.getClass().getClassLoader().getResource("icons/camera.png")).getScaledInstance(32, 32, Image.SCALE_DEFAULT )));
	  	capture.setToolTipText("Capture");
	 	capture.setBorder(BorderFactory.createRaisedBevelBorder());
	 	capture.setContentAreaFilled(true);
	 	capture.setLocation(0, 0);
	 	capture.setSize(50,50);
		capture.addActionListener(this);		
	 	totalGUI.add(capture);
	 	//Capture Button end
	 	
	 	//Doc Button start
	  	createDoc = new JButton(new ImageIcon( Toolkit.getDefaultToolkit().getImage(this.getClass().getClassLoader().getResource("icons/document.png")).getScaledInstance(32, 32, Image.SCALE_DEFAULT )));
	 	createDoc.setBorder(BorderFactory.createRaisedBevelBorder());
	 	createDoc.setContentAreaFilled(true);
	 	createDoc.setToolTipText("Create Document");
	 	createDoc.setLocation(51, 0);
	 	createDoc.setSize(50,50);
	 	createDoc.setEnabled(false);
	 	createDoc.addActionListener(this);
	  	totalGUI.add(createDoc);	 
		//Doc button end
	 	
	 	//Refresh button start
		refresh = new JButton(new ImageIcon( Toolkit.getDefaultToolkit().getImage(this.getClass().getClassLoader().getResource("icons/refresh.png")).getScaledInstance(32, 32, Image.SCALE_DEFAULT )));
	  	//	refresh = new JButton(new ImageIcon( Toolkit.getDefaultToolkit().getImage(".\\images\\refresh.png").getScaledInstance(32, 32, Image.SCALE_DEFAULT )));
	 	refresh.setToolTipText("Refresh");
	 	refresh.setBorder(BorderFactory.createRaisedBevelBorder());
	 	refresh.setContentAreaFilled(true);
	  	refresh.setLocation(101, 0);
	 	refresh.setSize(50,50);
	 	refresh.setEnabled(false);
	 	refresh.addActionListener(this);
		totalGUI.add(refresh);	
		//Refresh Button end
	  
	 	return totalGUI;
	    	
	    }
	    
	    private static void createAndShowGUI() {
	    	JFrame.setDefaultLookAndFeelDecorated(true);
	    	frame = new JFrame("Roboshot");
	    	frame.setAlwaysOnTop(true);
	    	WindowApp demo = new WindowApp();
	        frame.setContentPane(demo.createContentPane());
	        frame.getContentPane().add(new JProgressBar(0,100));
	        frame.setResizable(false);
	        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	        frame.setSize(161, 83);//(161, 83)
	        frame.setVisible(true);
	       
	      
	    }
	    
	 
	    
	    public void actionPerformed(ActionEvent e)  {
	    	
	        if(e.getSource() == capture)
	        {
	            //take screenshot
	        	
	        	takeScreenShot();
	        
	        	
	        }
	        else if(e.getSource() == createDoc)
	        {
	        	//generate Document
	        	frame.getContentPane().setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
	        	generateDocument();
	        	frame.getContentPane().setCursor(null);
	        	
	        }
	        else if (e.getSource() == refresh) {
	        	int option = JOptionPane.showConfirmDialog(
	        		    frame,
	        		    "Screenshots taken will be lost after refreshing.Are you Sure?",
	        		    "Confirm Refresh",
	        		    JOptionPane.YES_NO_OPTION);
	        	if (option == JOptionPane.OK_OPTION) {
	        		listImage = null;
	        		createDoc.setEnabled(false);
	        		refresh.setEnabled(false);
	        	}

	        }
	      
	      
	        
	        	
	     
		      
	       
	    }

	    private static P newImage( WordprocessingMLPackage wordMLPackage,
	    	    byte[] bytes,
	    	    String filenameHint, String altText, 
	    	    int id1, int id2, long cx) throws Exception {

	    	    BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

	    	    Inline inline = imagePart.createImageInline(filenameHint, altText,id1, id2, cx, false);

	    	    // Now add the inline in w:p/w:r/w:drawing
	    	    ObjectFactory factory = new ObjectFactory();
	    	    P  p = factory.createP();
	    	    R  run = factory.createR();             
	    	    p.getContent().add(run);       
	    	    Drawing drawing = factory.createDrawing();               
	    	    run.getContent().add(drawing);               
	    	    drawing.getAnchorOrInline().add(inline);
	    	   

	    	    return p;
	    	 }
	        
	   

    public static void main(String[] args) {
    	SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
	        
    
	}
    // Start of Helper methods
    private void takeScreenShot() {
    	
    	frame.setAlwaysOnTop(false);
    	frame.setVisible(false);
    	
    	byte [] bytes = null;
    	if (!createDoc.isEnabled()) {
        	createDoc.setEnabled(true);
        	refresh.setEnabled(true);
        	}
        	
            //take screenshot
        	if (listImage == null) {	        	
        	listImage = new LinkedList<byte[]>();	        	
        	}
        	Robot robot = null;
    		try {
    			robot = new Robot();
    		} catch (AWTException e1) {
    			// TODO Auto-generated catch block
    			e1.printStackTrace();
    		}
      
     
        try {
            Thread.sleep(50);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
        
        //Take active screenshot
        robot.keyPress(KeyEvent.VK_ALT);
        robot.delay(10);
        robot.keyPress(KeyEvent.VK_PRINTSCREEN);
        robot.delay(50);
        robot.keyRelease(KeyEvent.VK_PRINTSCREEN);
        robot.delay(10);
        robot.keyRelease(KeyEvent.VK_ALT);
        try {
            Thread.sleep(30);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
        	
    	   	 
        	 	Transferable transfer = null;				
        	    ByteArrayOutputStream baos = new ByteArrayOutputStream();
	        		
	        			try {
	        					    transfer = Toolkit.getDefaultToolkit().getSystemClipboard().getContents(null);	        			
	        					    RenderedImage image = null;  	        			   
								    image = (RenderedImage) transfer.getTransferData(DataFlavor.imageFlavor);
								    frame.setVisible(true);
							        frame.setAlwaysOnTop(true);
							    	ImageIO.write(image, "jpg", baos);
				        			
					        		 // commons-io.jar
					        		 bytes = IOUtils.toByteArray(new ByteArrayInputStream(baos.toByteArray()));
					        		
					        		 listImage.add(bytes);	
					        		 bytes=null;						
					        		 System.gc();
							
	        			}
	        			catch(IllegalStateException ee) {
	        				JOptionPane.showMessageDialog(this, "Clipboard Busy...Please try again...","Error",JOptionPane.ERROR_MESSAGE);
	        				
	        			}
	        			catch(HeadlessException ee1) {
	        				ee1.printStackTrace();
	        			}
	        			catch (UnsupportedFlavorException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
							JOptionPane.showMessageDialog(this, "Invalid format in Clipboard or No Active Window present...Please select a Window first & take the screenshot again","Error",JOptionPane.ERROR_MESSAGE);
							if (listImage.size() == 0) {
								createDoc.setEnabled(false);
								refresh.setEnabled(false);
							}
	        			}
        			    catch (IOException io) {
        			    	
        			    	io.printStackTrace();
        			    	
        			    }
	        			finally {
	        				 frame.setVisible(true);
						     frame.setAlwaysOnTop(true);
	        			}
	        		
			
	
    }
    
    private void generateDocument() {

        //generate doc
    	createDoc.setEnabled(false);
    	refresh.setEnabled(false);
    	capture.setEnabled(false);
    	String flName = "";
    	String flPath = "";
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new FileNameExtensionFilter("Microsoft Word (*.doc, *.docx)", "doc", "docx"));
        int option = chooser.showSaveDialog(this);
        if (option == JFileChooser.APPROVE_OPTION) {
        	//Progress Bar
        
        	
        	
        	
        	
          if(chooser.getSelectedFile()!=null)
                flPath =  chooser.getSelectedFile().getPath();
          		flName = chooser.getSelectedFile().getName();
          		  
	        
	        	WordprocessingMLPackage wordMLPackage = null;
	        	try {
					 wordMLPackage = WordprocessingMLPackage.createPackage();
				} catch (InvalidFormatException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	        	for (byte[] image : listImage) {
	        		 P p = null;	  	        		
	        		 String filenameHint = null;
	        		 String altText = null;
	        		 int id1 = 0;
	        		 int id2 = 1;
	        		
					try {
						p = newImage( wordMLPackage, image,filenameHint, altText,id1, id2,8000);
					} catch (Exception e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
	        		 // Now add our p to the document
	        		 wordMLPackage.getMainDocumentPart().addObject(p);
	        		
	        	}// end of For Loop...
	        	
	        	try {
	        		File f = null;
	        		if (flName.endsWith(".docx")) {
	        			f = new File(flPath);
	        			
	        		}
	        		else 
	        		{
	        			f = new File(flPath + ".docx");
	        		}
					wordMLPackage.save(f);
					JOptionPane.showMessageDialog(this, "Document generated successfully...");
					listImage = null;
		            wordMLPackage = null;
		           
					
				} 
	        	
	        	catch (Docx4JException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
					if ((e1.getCause() instanceof FileNotFoundException)) {
						JOptionPane.showMessageDialog(this, "File cannot be accessed as it may be used by some other process or has been deleted...Please try Again!!!","Error",JOptionPane.ERROR_MESSAGE);
						createDoc.setEnabled(true);
						refresh.setEnabled(true);
					}
					else 
					{
						JOptionPane.showMessageDialog(this, "Something went Wrong!!! :(","Error",JOptionPane.ERROR_MESSAGE);
					}
					createDoc.setEnabled(true);
					refresh.setEnabled(true);
					
				}
	        	finally {
	        		 capture.setEnabled(true);
	        	}
	        	
	        
	       
	        	              		
        }
        else {
        	createDoc.setEnabled(true);
        	refresh.setEnabled(true);
        	capture.setEnabled(true);
        }
      
        
    
    }
    
    //End Of Helper Methods
    
}
	


