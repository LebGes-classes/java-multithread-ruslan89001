import java.util.concurrent.*;
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeTaskManager {
    private static final int WORK_HOURS = 8;

    public static void main(String[] args) {
        ExecutorService executor = Executors.newFixedThreadPool(5); // 5 сотрудников

        try (Workbook workbook = WorkbookFactory.create(new File("employee_data.xlsx"))) {
            Sheet sheet = workbook.getSheetAt(0);

            // Загрузка данных сотрудников из xlsx файла
            int[] taskHours = loadEmployeeData(sheet);

            for (int i = 0; i < taskHours.length; i++) { // Задачи для каждого сотрудника
                final int hours = taskHours[i];
                executor.submit(() -> {
                    try {
                        manageTask(hours);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                });
            }

            executor.shutdown();
            try {
                if (!executor.awaitTermination(60, TimeUnit.SECONDS)) {
                    executor.shutdownNow();
                }
            } catch (InterruptedException ex) {
                executor.shutdownNow();
                Thread.currentThread().interrupt();
            }

            // Сохранение данных в конце рабочего дня
            try (Workbook newWorkbook = new XSSFWorkbook()) {
                Sheet newSheet = newWorkbook.createSheet("Updated Data");
                saveEmployeeData(newWorkbook, newSheet, sheet, "employee_data_updated.xlsx");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static int[] loadEmployeeData(Sheet sheet) {
        int[] taskHours = new int[sheet.getLastRowNum()];
        // Пропустить заголовок
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String employeeName = row.getCell(0).getStringCellValue();
            taskHours[i - 1] = (int) row.getCell(1).getNumericCellValue();

            // Загрузить данные сотрудника и задачи
            System.out.println("Employee: " + employeeName + ", Task Hours: " + taskHours[i - 1]);
        }
        return taskHours;
    }

    private static void manageTask(int taskHours) throws InterruptedException {
        int remainingHours = taskHours;
        int day = 1;
        while (remainingHours > 0) {
            int workHours = Math.min(remainingHours, WORK_HOURS);
            System.out.println("Day " + day + ": Working for " + workHours + " hours");
            remainingHours -= workHours;
            day++;
        }
    }

    private static void saveEmployeeData(Workbook newWorkbook, Sheet newSheet, Sheet oldSheet, String fileName) {
        // Записать данные сотрудника и статистику в файл Excel
        int rowNum = 0;
        for (int i = 0; i < oldSheet.getLastRowNum(); i++) { // Для каждого сотрудника
            Row oldRow = oldSheet.getRow(i + 1);
            String employeeName = oldRow.getCell(0).getStringCellValue();
            int taskHours = (int) oldRow.getCell(1).getNumericCellValue();
            int workHours = Math.min(taskHours, WORK_HOURS);
            int totalWorkHours = WORK_HOURS * (taskHours / WORK_HOURS + (taskHours % WORK_HOURS > 0 ? 1 : 0)); // Общее время работы
            int idleHours = totalWorkHours - taskHours; // Время простоя
            double efficiency = (double) taskHours / totalWorkHours; // Эффективность

            Row newRow = newSheet.createRow(rowNum++);
            newRow.createCell(0).setCellValue(employeeName);
            newRow.createCell(1).setCellValue("Task Hours: " + taskHours);
            newRow.createCell(2).setCellValue("Work Hours: " + workHours);
            newRow.createCell(3).setCellValue("Total Work Hours: " + totalWorkHours);
            newRow.createCell(4).setCellValue("Idle Hours: " + idleHours);
            newRow.createCell(5).setCellValue("Efficiency: " + efficiency);
        }

        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            newWorkbook.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
