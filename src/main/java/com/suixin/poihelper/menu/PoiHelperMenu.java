package com.suixin.poihelper.menu;

import com.intellij.openapi.actionSystem.AnAction;
import com.intellij.openapi.actionSystem.AnActionEvent;
import com.intellij.openapi.fileEditor.FileEditorManager;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.vfs.VirtualFile;
import com.intellij.openapi.vfs.VirtualFileManager;
import com.suixin.poihelper.entity.LineEntity;
import org.jsoup.internal.StringUtil;

import javax.swing.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

public class PoiHelperMenu extends AnAction {
    @Override
    public void actionPerformed(AnActionEvent e) {
        // 获取当前文件名称
        Project project = e.getProject(); // 获取当前的 Project 对象
        FileEditorManager fileEditorManager = FileEditorManager.getInstance(project);
        VirtualFile[] selectedFiles = fileEditorManager.getSelectedFiles();
        String path = selectedFiles[0].getPath();
        BufferedReader reader = null;
        try {
            reader = new BufferedReader(new FileReader(path, StandardCharsets.UTF_8));
        } catch (IOException ex) {
            throw new RuntimeException(ex);
        }
        List<String> lines = reader.lines().collect(Collectors.toList());
        List<LineEntity> numberList = new ArrayList<>();
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            String line2 = lines.get(i);
//            @Excel(name = "员工编号",orderNum = "1")
            String singleLineComments = "";
            if (!StringUtil.isBlank(line2)) {
                String trim = line2.trim();
                if (trim.length() > 1) {
                    singleLineComments = trim.substring(0, 2);
                }
            }
            if (line.contains("@Excel") && !singleLineComments.equals("//")) {
                String[] split = line.split(",");
                for (String s : split) {
                    if (s.contains("orderNum")) {
                        String[] orderNumSplit = s.split("=");
                        String s1 = orderNumSplit[1];
                        int firstIndex = s1.indexOf("\"");
                        int lastIndex = s1.lastIndexOf("\"");
                        Integer currentLineNumber = Integer.valueOf(s1.substring(firstIndex+1, lastIndex).trim());
                        LineEntity lineEntity = new LineEntity();
                        lineEntity.setLine(i);
                        lineEntity.setValue(currentLineNumber);
                        numberList.add(lineEntity);
                    }
                }
            }
        }
        if (numberList.size() == 0) {
            return;
        }
        bubbleSort(numberList);
        List<LineEntity> newList = new ArrayList<>();
        sortList(numberList,newList);
        newList = sortList2(newList);

        Map<Integer,Integer> exportRows = new HashMap<>();
        for (LineEntity lineEntity : newList) {
            exportRows.put(lineEntity.getLine(), lineEntity.getValue());
        }
        List<String> newContent = new ArrayList<>();
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            String line2 = lines.get(i);
//            @Excel(name = "员工编号",orderNum = "1")
            String singleLineComments = "";
            if (!StringUtil.isBlank(line2)) {
                String trim = line2.trim();
                if (trim.length() > 1) {
                    singleLineComments = trim.substring(0, 2);
                }
            }
            if (line.contains("@Excel") && !singleLineComments.equals("//")) {
                String newLine = "";
                String[] split = line.split(",");
                for (String s : split) {
                    if (s.contains("orderNum")) {
                        String[] orderNumSplit = s.split("=");
                        String s1 = orderNumSplit[1];
                        int firstIndex = s1.indexOf("\"");
                        Integer currentLineNumber = exportRows.get(i);
                        if (s.contains(")")) {
                            newLine = newLine + orderNumSplit[0] + "=" +s1.substring(0, firstIndex+1) + currentLineNumber+"\"),";
                        }else{
                            newLine = newLine + orderNumSplit[0] + "=" + s1.substring(0, firstIndex+1) + currentLineNumber+"\",";
                        }
                    }else {
                        newLine = newLine + s + ",";
                    }
                }
                newLine = newLine.substring(0, newLine.length() - 1);
                newContent.add(newLine);
            }else {
                newContent.add(line);
            }
        }
        try {
            saveToXMLFile(newContent,path);
            VirtualFileManager.getInstance().syncRefresh();
            JOptionPane.showMessageDialog(null, "排序完成", "提示", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException ex) {
            throw new RuntimeException(ex);
        }
    }



    // 冒泡排序算法
    public static void bubbleSort(List<LineEntity> numberList) {
        int n = numberList.size();
        for (int i = 0; i < n-1; i++) {
            for (int j = 0; j < n-i-1; j++) {
                // 如果当前元素大于下一个元素，交换它们
                if (numberList.get(j).getValue() > numberList.get(j+1).getValue()) {
                    LineEntity temp = numberList.get(j);
                    numberList.set(j,numberList.get(j+1));
                    numberList.set(j+1, temp);
                }
            }
        }
    }

    static void sortList(List<LineEntity> list, List<LineEntity> newList){
        Integer repeatNumber = null;
        Integer repeatCount = 0;
        for (LineEntity i : list) {
            if (repeatNumber == null) {
                repeatNumber = i.getValue();
                newList.add(i);
                continue;
            }
            if (repeatNumber.equals(i.getValue())) {
                repeatCount++;
                i.setValue(i.getValue() + repeatCount);
                newList.add(i);
            }else {
                repeatNumber = i.getValue();
                i.setValue(i.getValue() + repeatCount);
                newList.add(i);
            }
        }
        Set<LineEntity> uniqueSet = new LinkedHashSet<>(newList);
        if (uniqueSet.size() != newList.size()) {
            list.clear();
            list.addAll(newList);
            newList.clear();
            sortList(list,newList);
        }
    }
    static List<LineEntity> sortList2(List<LineEntity> list){
        List<LineEntity> newList = new ArrayList<>();
        newList.addAll(list);
        for (int i = 1; i < list.size(); i++) {
            Integer number = list.get(i).getValue();
            Integer previousNumber = newList.get(i-1).getValue();
            if (number - previousNumber > 1) {
                LineEntity lineEntity = newList.get(i);
                lineEntity.setValue(previousNumber+1);
                newList.set(i, lineEntity);
            }
        }
        return newList;
    }

    public static void saveToXMLFile(List<String> newContent, String newFileName) throws IOException {
        File file = new File(newFileName);
        if (!file.exists()) {
            file.createNewFile();
        }
        String xml = "";
        for (String text : newContent) {
            xml = xml + text + "\n";
        }

        FileOutputStream outStream = null;
        OutputStreamWriter writer = null;
        try {
            outStream = new FileOutputStream(file);
            writer = new OutputStreamWriter(outStream, "UTF-8");
            writer.write(xml);
        } finally {
            if (writer != null) {
                writer.close();
            }
            if (outStream != null) {
                outStream.close();
            }
        }
    }
}
