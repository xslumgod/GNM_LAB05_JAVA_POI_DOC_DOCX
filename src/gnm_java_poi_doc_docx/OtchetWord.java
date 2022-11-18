package gnm_java_poi_doc_docx;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class OtchetWord extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread {

        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            
            // Чтение из файла-шаблона в переменную doc
            HWPFDocument doc = null;
            try (FileInputStream fis = new FileInputStream(dir + "otchet_template.doc")) {
                doc = new HWPFDocument(fis);
                fis.close();
            } catch (Exception ex) {
                System.err.println("Error template!");
            }

            // Замена в переменной doc данных
            try {
                doc.getRange().replaceText("$Выручка", jTextField_1.getText());
                doc.getRange().replaceText("$Сибестоимость", jTextField_2.getText());
                doc.getRange().replaceText("$Убыток", jTextField_3.getText());
                doc.getRange().replaceText("$КоммерчискиеРасходы", jTextField_4.getText());
                doc.getRange().replaceText("$УправленческиеРасходы", jTextField_5.getText());
                doc.getRange().replaceText("$ПрибыльОтПродаж", jTextField_6.getText());
                doc.getRange().replaceText("$выручка", jTextField_7.getText());
                doc.getRange().replaceText("$сибестоимость", jTextField_8.getText());
                doc.getRange().replaceText("$убыток", jTextField_9.getText());
                doc.getRange().replaceText("$коммерчискиеРасходы", jTextField_10.getText());
                doc.getRange().replaceText("$управленческиеРасходы", jTextField_11.getText());
                doc.getRange().replaceText("$прибыльОтПродаж", jTextField_12.getText());
                doc.getRange().replaceText("$Организация", jTextField_13.getText());
                doc.getRange().replaceText("$ИИН", jTextField_14.getText());
            } catch (Exception ex) {
                System.err.println("Error replaceText!");
            }

            // Сохранение переменной doc в новый файл
            try (FileOutputStream fos = new FileOutputStream(dir + "otchet.doc")) {
                doc.write(fos);
                fos.close();
                // Открытие файла внешней программой
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "otchet.doc").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "otchet.doc"));
                }
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }
    
    class TThread2 extends Thread {

        
        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            
            // Чтение из файла-шаблона в переменную docx
            XWPFDocument doc = null;
            
           try (FileInputStream fis = new FileInputStream(dir + "otchet_template.docx")) {
                doc = new XWPFDocument(fis);
                fis.close();
            } catch (Exception ex) {
                System.err.println("Error template!");
            }

            // Замена в переменной doc данных
            
            for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains("$Выручка")) {
                                text = text.replace(text, jTextField_1.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$Сибестоимость")) {
                                text = text.replace(text, jTextField_2.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$Убыток")) {
                                text = text.replace(text, jTextField_3.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$КоммерчискиеРасходы")) {
                                text = text.replace(text, jTextField_4.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$УправленческиеРасходы")) {
                                text = text.replace(text, jTextField_5.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$ПрибыльОтПродаж")) {
                                text = text.replace(text, jTextField_6.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$выручка")) {
                                text = text.replace(text, jTextField_7.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$сибестоимость")) {
                                text = text.replace(text, jTextField_8.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$убыток")) {
                                text = text.replace(text, jTextField_9.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$коммерчискиеРасходы")) {
                                text = text.replace(text, jTextField_10.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$управленческиеРасходы")) {
                                text = text.replace(text, jTextField_11.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$прибыльОтПродаж")) {
                                text = text.replace(text, jTextField_12.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$Организация")) {
                                text = text.replace(text, jTextField_13.getText());
                                r.setText(text, 0);
                            }
                            if (text != null && text.contains("$ИИН")) {
                                text = text.replace(text, jTextField_14.getText());
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }

            // Сохранение переменной docx в новый файл
            
                        try (FileOutputStream fos = new FileOutputStream(dir + "otchet.docx")) {
                doc.write(fos);
                fos.close();
                // Открытие файла внешней программой
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "otchet.docx").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "otchet.docx"));
                }
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
        
    }

    public OtchetWord() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton_Save = new javax.swing.JButton();
        jButton_Save1 = new javax.swing.JButton();
        jTextField_1 = new javax.swing.JTextField();
        jTextField_2 = new javax.swing.JTextField();
        jTextField_3 = new javax.swing.JTextField();
        jTextField_4 = new javax.swing.JTextField();
        jTextField_5 = new javax.swing.JTextField();
        jTextField_6 = new javax.swing.JTextField();
        jTextField_7 = new javax.swing.JTextField();
        jTextField_8 = new javax.swing.JTextField();
        jTextField_9 = new javax.swing.JTextField();
        jTextField_10 = new javax.swing.JTextField();
        jTextField_11 = new javax.swing.JTextField();
        jTextField_12 = new javax.swing.JTextField();
        jTextField_13 = new javax.swing.JTextField();
        jTextField_14 = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Отчет в MS Word");
        setResizable(false);
        setSize(new java.awt.Dimension(0, 0));
        getContentPane().setLayout(null);

        jButton_Save.setText("в WORD");
        jButton_Save.setToolTipText("");
        jButton_Save.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_SaveActionPerformed(evt);
            }
        });
        getContentPane().add(jButton_Save);
        jButton_Save.setBounds(930, 50, 80, 23);

        jButton_Save1.setText("в WORDX");
        jButton_Save1.setToolTipText("");
        jButton_Save1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_Save1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton_Save1);
        jButton_Save1.setBounds(1020, 50, 100, 23);

        jTextField_1.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_1);
        jTextField_1.setBounds(670, 540, 263, 30);

        jTextField_2.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_2);
        jTextField_2.setBounds(670, 578, 263, 30);

        jTextField_3.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_3);
        jTextField_3.setBounds(670, 620, 260, 30);

        jTextField_4.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_4.setPreferredSize(new java.awt.Dimension(55, 20));
        getContentPane().add(jTextField_4);
        jTextField_4.setBounds(670, 670, 260, 30);

        jTextField_5.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_5ActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_5);
        jTextField_5.setBounds(670, 700, 260, 30);

        jTextField_6.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_6);
        jTextField_6.setBounds(670, 750, 260, 30);

        jTextField_7.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_7);
        jTextField_7.setBounds(940, 540, 200, 30);

        jTextField_8.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_8);
        jTextField_8.setBounds(940, 578, 200, 30);

        jTextField_9.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_9);
        jTextField_9.setBounds(940, 620, 200, 30);

        jTextField_10.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_10);
        jTextField_10.setBounds(940, 670, 200, 30);

        jTextField_11.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_11);
        jTextField_11.setBounds(940, 700, 200, 30);

        jTextField_12.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_12);
        jTextField_12.setBounds(940, 750, 200, 30);

        jTextField_13.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_13);
        jTextField_13.setBounds(190, 190, 260, 30);

        jTextField_14.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_14);
        jTextField_14.setBounds(940, 220, 200, 25);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/gnm_java_poi_doc_docx/otchet.png"))); // NOI18N
        getContentPane().add(jLabel1);
        jLabel1.setBounds(10, 10, 1150, 800);

        setSize(new java.awt.Dimension(1201, 840));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton_SaveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_SaveActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton_SaveActionPerformed

    private void jTextField_5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_5ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_5ActionPerformed

    private void jButton_Save1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_Save1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread2().start();
    }//GEN-LAST:event_jButton_Save1ActionPerformed

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
            java.util.logging.Logger.getLogger(OtchetWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(OtchetWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(OtchetWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(OtchetWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new OtchetWord().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton_Save;
    private javax.swing.JButton jButton_Save1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JTextField jTextField_1;
    private javax.swing.JTextField jTextField_10;
    private javax.swing.JTextField jTextField_11;
    private javax.swing.JTextField jTextField_12;
    private javax.swing.JTextField jTextField_13;
    private javax.swing.JTextField jTextField_14;
    private javax.swing.JTextField jTextField_2;
    private javax.swing.JTextField jTextField_3;
    private javax.swing.JTextField jTextField_4;
    private javax.swing.JTextField jTextField_5;
    private javax.swing.JTextField jTextField_6;
    private javax.swing.JTextField jTextField_7;
    private javax.swing.JTextField jTextField_8;
    private javax.swing.JTextField jTextField_9;
    // End of variables declaration//GEN-END:variables
}
