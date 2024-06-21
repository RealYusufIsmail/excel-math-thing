package io.github.realyusufismail;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.data.general.DefaultPieDataset;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.List;

public class ScoreChartApp extends JFrame {
    private final JTextField filePathField;
    private final JTextField columnField;
    private final JTextField maxScoreField;
    private final JTextField titleField;

    public ScoreChartApp() {
        setTitle("Score Chart Generator");
        setSize(600, 200);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new GridLayout(5, 3));

        JLabel filePathLabel = new JLabel("Excel File Path:");
        filePathField = new JTextField();
        JButton browseButton = new JButton("Browse");
        browseButton.addActionListener(this::browseFile);

        JLabel columnLabel = new JLabel("Score Column Name:");
        columnField = new JTextField();

        JLabel maxScoreLabel = new JLabel("Max Score:");
        maxScoreField = new JTextField();

        JLabel titleLabel = new JLabel("Chart Title:");
        titleField = new JTextField();

        JButton generateButton = new JButton("Generate Chart");
        generateButton.addActionListener(this::generateChart);

        add(filePathLabel);
        add(filePathField);
        add(browseButton);
        add(columnLabel);
        add(columnField);
        add(new JLabel()); // Empty cell
        add(maxScoreLabel);
        add(maxScoreField);
        add(new JLabel()); // Empty cell
        add(titleLabel);
        add(titleField);
        add(new JLabel()); // Empty cell
        add(generateButton);

        setVisible(true);
    }

    private void browseFile(ActionEvent event) {
        JFileChooser fileChooser = new JFileChooser();
        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            filePathField.setText(selectedFile.getAbsolutePath());
        }
    }

    private void generateChart(ActionEvent event) {
        String filePath = filePathField.getText();
        String scoreColumn = columnField.getText();
        String maxScoreText = maxScoreField.getText();
        String chartTitle = titleField.getText();

        if (filePath.isEmpty() || scoreColumn.isEmpty() || maxScoreText.isEmpty() || chartTitle.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please provide file path, score column name, max score, and chart title.", "Input Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        int maxScore;
        try {
            maxScore = Integer.parseInt(maxScoreText);
        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(this, "Max score must be a valid integer.", "Input Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        try (FileInputStream file = new FileInputStream(new File(filePath))) {
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            Map<Integer, Integer> scoreCounts = new HashMap<>();
            for (int i = 1; i <= maxScore; i++) {
                scoreCounts.put(i, 0);
            }

            int numberOfParticipants = 0;

            Iterator<Row> rowIterator = sheet.iterator();
            int scoreColumnIndex = -1;

            // Find the index of the target column
            if (rowIterator.hasNext()) {
                Row headerRow = rowIterator.next();
                for (Cell cell : headerRow) {
                    if (cell.getStringCellValue().equalsIgnoreCase(scoreColumn)) {
                        scoreColumnIndex = cell.getColumnIndex();
                        break;
                    }
                }
            }

            if (scoreColumnIndex == -1) {
                JOptionPane.showMessageDialog(this, "Column '" + scoreColumn + "' not found in the Excel file.", "Input Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            // Iterate over rows to extract data until the column is blank
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(scoreColumnIndex);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    continue;
                }

                int score;
                if (cell.getCellType() == CellType.NUMERIC) {
                    score = (int) cell.getNumericCellValue();
                } else if (cell.getCellType() == CellType.STRING) {
                    try {
                        score = Integer.parseInt(cell.getStringCellValue());
                    } catch (NumberFormatException e) {
                        continue;
                    }
                } else {
                    continue;
                }

                if (score >= 1 && score <= maxScore) {
                    scoreCounts.put(score, scoreCounts.get(score) + 1);
                    numberOfParticipants++;
                }
            }

            List<Float> percentages = new ArrayList<>();
            for (int i = 1; i <= maxScore; i++) {
                percentages.add(calculatePercentage(numberOfParticipants, scoreCounts.get(i)));
            }

            List<Float> cleanPercentages = roundPercentages(percentages);

            DefaultPieDataset dataset = new DefaultPieDataset();
            for (int i = 1; i <= maxScore; i++) {
                float percentage = cleanPercentages.get(i - 1);
                if (percentage > 0) {
                    dataset.setValue(i + " (" + String.format("%.2f", percentage) + "%)", percentage);
                }
            }

            JFreeChart pieChart = ChartFactory.createPieChart(chartTitle, dataset, true, true, false);
            ChartPanel pieChartPanel = new ChartPanel(pieChart);

            JFrame pieChartFrame = new JFrame("Pie Chart");
            pieChartFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
            pieChartFrame.add(pieChartPanel);
            pieChartFrame.pack();
            pieChartFrame.setVisible(true);

        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Error: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
        }
    }

    private float calculatePercentage(int total, int count) {
        return (float) count / total * 100;
    }

    private List<Float> roundPercentages(List<Float> percentages) {
        List<Float> roundedPercentages = new ArrayList<>();
        float total = 0;

        for (float percentage : percentages) {
            float rounded = Math.round(percentage);
            roundedPercentages.add(rounded);
            total += rounded;
        }

        if (total != 100) {
            float difference = 100 - total;
            int minIndex = 0;
            for (int i = 1; i < roundedPercentages.size(); i++) {
                if (roundedPercentages.get(i) < roundedPercentages.get(minIndex)) {
                    minIndex = i;
                }
            }
            roundedPercentages.set(minIndex, roundedPercentages.get(minIndex) + difference);
        }

        return roundedPercentages;
    }
}
