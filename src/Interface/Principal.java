
package Interface;

import java.awt.Dimension;
import java.io.File;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.table.DefaultTableModel;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class Principal extends javax.swing.JFrame {


File file;
Workbook workbook;




public Principal() {
    initComponents();
}


@SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jtlocal = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jttitulo = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tb1 = new javax.swing.JTable();
        jbpreencher = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(null);
        getContentPane().add(jtlocal);
        jtlocal.setBounds(60, 40, 400, 40);

        jButton1.setText("XLS");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(490, 40, 130, 40);
        getContentPane().add(jttitulo);
        jttitulo.setBounds(60, 120, 400, 40);

        jLabel1.setText("Titulo da Planilha");
        getContentPane().add(jLabel1);
        jLabel1.setBounds(70, 100, 200, 16);

        tb1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Item", "Descrição"
            }
        ));
        jScrollPane1.setViewportView(tb1);

        getContentPane().add(jScrollPane1);
        jScrollPane1.setBounds(60, 180, 560, 270);

        jbpreencher.setText("Preencher");
        jbpreencher.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbpreencherActionPerformed(evt);
            }
        });
        getContentPane().add(jbpreencher);
        jbpreencher.setBounds(490, 120, 130, 40);

        setSize(new java.awt.Dimension(687, 521));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed

        //// inicio do codigo
        JFileChooser fc = new JFileChooser();
         fc.setPreferredSize( new Dimension(750, 400) );
         
         fc.setFileSelectionMode(JFileChooser.FILES_ONLY );
        
         int resultado = fc.showOpenDialog(this);
               
                if ( resultado == JFileChooser.CANCEL_OPTION ) {
                    //para cancelar a janela se necessário
                }
                else{
                 file = fc.getSelectedFile();                
                            
                 jtlocal.setText(file.toString().trim());
                    
                }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jbpreencherActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jbpreencherActionPerformed

    try {
        workbook = Workbook.getWorkbook(new File(jtlocal.getText().trim()));
        
        
    } catch (IOException | BiffException ex) {
        Logger.getLogger(Principal.class.getName()).log(Level.SEVERE, null, ex);
    }
    
                Sheet sheet = workbook.getSheet(0);
                             
                Cell c0 = sheet.getCell(0, 0);
                String titulo = c0.getContents();               
                
               
               jttitulo.setText(titulo);  
               
               
               int linhas = sheet.getRows();
           
               ////inicio do for
            for(int i = 2; i < linhas; i++){

                    Cell ca = sheet.getCell(0, i);
                    Cell cb = sheet.getCell(1, i);

                  
                    String item = ca.getContents();
                    String desc = cb.getContents();
              
               
                        DefaultTableModel mp = (DefaultTableModel) tb1.getModel();                        
                        mp.addRow(new String[]{item,desc});

                    

             }   //////// fim do for

                workbook.close();
   

    }//GEN-LAST:event_jbpreencherActionPerformed

public static void main(String args[]) {
    /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
     * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
     */
    try {
        for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
            if ("Nimbus".equals(info.getName())) {
                javax.swing.UIManager.setLookAndFeel(info.getClassName());
                break;
            }
        }
    } catch (ClassNotFoundException ex) {
        java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    } catch (InstantiationException ex) {
        java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    } catch (IllegalAccessException ex) {
        java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    } catch (javax.swing.UnsupportedLookAndFeelException ex) {
        java.util.logging.Logger.getLogger(Principal.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    }
        //</editor-fold>

    /* Create and display the form */
    java.awt.EventQueue.invokeLater(new Runnable() {
    public void run() {
        new Principal().setVisible(true);
    }
    });
}

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JButton jbpreencher;
    private javax.swing.JTextField jtlocal;
    private javax.swing.JTextField jttitulo;
    private javax.swing.JTable tb1;
    // End of variables declaration//GEN-END:variables
}
