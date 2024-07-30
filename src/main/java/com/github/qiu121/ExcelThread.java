package com.github.qiu121;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.LinkedBlockingQueue;


// ExcelThread类继承自Thread类，用于处理Excel文件的读取
public class ExcelThread extends Thread {
    // 输入文件的路径
    private final String inputFilename;
    // 使用BlockingQueue来存储从Excel文件中读取的行，以便在多线程环境中进行线程安全的操作
    private static BlockingQueue<Row> queue = new LinkedBlockingQueue<>();

    // 构造函数，接收一个输入文件的路径
    public ExcelThread(String inputFilename) {
        this.inputFilename = inputFilename;
    }

    // 重写run方法，这是线程启动后执行的方法
    @Override
    public void run() {
        // 这是一个无限循环，除非程序停止，否则线程会一直运行
        while (true) {
            try {
                // 创建一个文件输入流，用于读取Excel文件
                FileInputStream fis = new FileInputStream(inputFilename);
                // 创建一个工作簿对象，用于操作Excel文件
                Workbook workbook = new XSSFWorkbook(fis);
                // 获取工作簿的第一个工作表
                Sheet sheet = workbook.getSheetAt(0);

                // 遍历工作表的每一行
                for (Row row : sheet) {
                    // 如果是第一行（表头），则跳过
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    // 遍历行的每一个单元格
                    for (Cell cell : row) {
                        // 如果单元格存在且类型为数值型，则将其值乘以2
                        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                            double value = cell.getNumericCellValue();
                            cell.setCellValue(value * 2);
                        }
                    }
                    // 将处理过的行添加到队列中
                    queue.put(row);
                }

                // 关闭文件输入流
                fis.close();
                // 让线程休眠2秒，这样可以模拟读取文件的过程需要一些时间
                Thread.sleep(2000);
            } catch (IOException | InterruptedException e) {
                // 如果捕获到异常，则打印异常栈信息
                e.printStackTrace();
            }
        }
    }

    // 主方法，程序的入口点
    public static void main(String[] args) {
        // 创建两个ExcelThread线程，分别处理两个输入文件
        Thread thread1 = new ExcelThread("src/main/resources/x1.xlsx");
        Thread thread2 = new ExcelThread("src/main/resources/x2.xlsx");

        // 启动两个线程
        thread1.start();
        thread2.start();

        // 创建一个单线程的线程池，用于执行写入操作
        ExecutorService executor = Executors.newSingleThreadExecutor();
        executor.submit(() -> {
            // 这是一个无限循环，除非程序停止，否则线程会一直运行
            while (true) {
                try {
                    // 创建一个File对象，表示输出文件
                    File file = new File("src/main/resources/x3.xlsx");
                    Workbook workbookOut;
                    // 如果文件不存在，则创建新文件，并写入表头
                    if (!file.exists()) {
                        file.createNewFile();
                        workbookOut = new XSSFWorkbook();
                        Sheet sheetOut = workbookOut.createSheet();

                        Row headerRow = sheetOut.createRow(0);
                        headerRow.createCell(0).setCellValue("序号");
                        headerRow.createCell(1).setCellValue("平方");
                        headerRow.createCell(2).setCellValue("价格");
                        headerRow.createCell(3).setCellValue("每平价格");
                        headerRow.createCell(4).setCellValue("时间戳");
                    } else {
                        // 如果文件已存在，则打开文件
                        FileInputStream fis = new FileInputStream(file);
                        workbookOut = new XSSFWorkbook(fis);
                        fis.close();
                    }

                    // 获取工作簿的第一个工作表
                    Sheet sheetOut = workbookOut.getSheetAt(0);
                    // 如果工作表不存在，则创建新工作表
                    if (sheetOut == null) {
                        sheetOut = workbookOut.createSheet();
                    }

                    // 从队列中取出一行
                    Row row = queue.take();


                    // 在输出文件的工作表中创建新行
                    Row newRow = sheetOut.createRow(sheetOut.getLastRowNum() + 1);
                    // 创建一个新的样式
                    CellStyle style = workbookOut.createCellStyle();
                    // 设置样式的水平对齐方式为居中
                    style.setAlignment(HorizontalAlignment.CENTER);
                    // 设置样式的垂直对齐方式为居中
                    style.setVerticalAlignment(VerticalAlignment.CENTER);


                    // 遍历输入行的每一个单元格
                    for (Cell cell : row) {
                        // 在新行中创建新单元格，并复制输入单元格的值
                        Cell newCell = newRow.createCell(cell.getColumnIndex(), cell.getCellType());
                        // 将样式应用到新单元格
                        newCell.setCellStyle(style);
                        switch (cell.getCellType()) {
                            case STRING:
                                newCell.setCellValue(cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                newCell.setCellValue(cell.getNumericCellValue());
                                break;
                            default:
                                break;
                        }
                    }

                    // 在新行的第五列（索引从0开始）创建新单元格，用于存储时间戳
                    Cell newCell = newRow.createCell(4);
                    // 将样式应用到新单元格
                    newCell.setCellStyle(style);


                    // 将当前日期时间格式化为字符串，并存储到新单元格中
                    DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
                    newCell.setCellValue(LocalDateTime.now().format(dateTimeFormatter));

                    // 设置第五列（索引从0开始）的宽度为3个默认格子宽度,8是默认字符宽度
                    sheetOut.setColumnWidth(4, 3 * 256 * 8);

                    // 创建一个文件输出流，用于写入Excel文件
                    FileOutputStream fos = new FileOutputStream("src/main/resources/x3.xlsx");
                    // 将工作簿的内容写入文件
                    workbookOut.write(fos);
                    // 关闭文件输出流
                    fos.close();
                } catch (InterruptedException | IOException e) {
                    // 如果捕获到异常，则打印异常栈信息
                    e.printStackTrace();
                }
            }
        });
    }
}
