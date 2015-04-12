package Ict4d;

import javax.swing.*;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import java.awt.*;
 
public class GUI extends JPanel implements ChangeListener {
    private JTextField field;
    private JTextField price;
    private JTextField propety;
    public static double calculation = 0;
    static Propety[] PropetyRow  = new Propety[500];
     public GUI() {
        initializeUI();
    }
 
    private void initializeUI() {
        setLayout(new BorderLayout());
        setPreferredSize(new Dimension(800, 800));
        JComboBox<String> propetyList = new JComboBox<String>();
        JSlider slider = new JSlider(JSlider.HORIZONTAL, 0, 10, 0);
        
        slider.setPaintTicks(true);
        slider.setPaintLabels(true);
        slider.setMinorTickSpacing(1);
        slider.setMajorTickSpacing(1);
        slider.addChangeListener(this);
 
        JLabel label = new JLabel("Rent:");
        price = new JTextField(5);
        
 
        JPanel panel = new JPanel();
        panel.setLayout(new FlowLayout());
        panel.add(label);
   
        panel.add(price);
 
        add(slider, BorderLayout.PAGE_END);
  
        add(panel, BorderLayout.EAST);
    }
 
    public void stateChanged(ChangeEvent e) {
        JSlider slider = (JSlider) e.getSource();
 
        //
        // Get the selection value of JSlider
        Estate_Prediction.calculate(PropetyRow, slider.getValue(),1);
        Estate_Prediction.calculation = slider.getValue();
        price.setText(String.valueOf(calculation));
        
        
        
       
    }
 
    public static void showFrame(Propety [] propety) {
        JPanel panel = new GUI();
        PropetyRow = propety;
        panel.setOpaque(true);
        JFrame frame = new JFrame();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setTitle("Estate_Prediction");
        frame.setContentPane(panel);
        frame.pack();
        frame.setVisible(true);
    }
 
  
        
    }


