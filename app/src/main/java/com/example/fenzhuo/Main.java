package com.example.fenzhuo;

import android.util.Log;

import com.carl.excel.Entry;
import com.carl.excel.ExcelUtil;
import com.carl.excel.config.ConfigUtil;
import com.carl.excel.config.ConfigXlsx2Xml;
import com.google.gson.Gson;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class Main {
    public static final int NUMBER_OF_PER_TABLE = 10;
    public static final int OLD_PRIORITY = 1;
    public static final int SEX_PRIORITY = 1;
    public static final int SECTION_PRIORITY = 1;
    static int tableCount;
    static ArrayList<Table> tableArrayList;

    public static void main(String[] args) {
        go();
    }

    public static void go() {
        ArrayList<Staff> allStaffArrayList = initData();
        Collections.shuffle(allStaffArrayList);
        tableCount = (int) Math.ceil(((float) allStaffArrayList.size()) / NUMBER_OF_PER_TABLE);
        tableArrayList = new ArrayList<>();
        for (int i = 0; i < tableCount; i++) {
            Table table = new Table();
            table.tableId = i + "";
            tableArrayList.add(table);
        }
        chooseTable(allStaffArrayList);
    }

    public static ArrayList<Staff> initData() {
        ArrayList<Staff> temp = new ArrayList<>();
        String configPath = ConfigUtil.PATH_CONFIG_XLSX2XML;
        List<ConfigXlsx2Xml> configList = ConfigXlsx2Xml.parse(configPath);
        if (configList == null || configList.size() <= 0) {
            System.out.println("Xlsx2XmlModel=null");
        }

        try {
            for (ConfigXlsx2Xml config : configList) {
                ArrayList<Staff> entryList = ExcelUtil.ReadExcel(config.getXlsxPath(),
                        config.getSheetName(), config.getName(), config.getSex(), config.getXinlao(), config.getBumen());
                return entryList;
            }
        } catch (InvalidFormatException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return temp;
    }

    /**
     * 每个选中选择合适餐桌入座
     *
     * @param staffArrayList
     */
    private static void chooseTable(ArrayList<Staff> staffArrayList) {
        ArrayList<Staff> noSitStaffList = getNoSitStaffList(staffArrayList);
        for (int i = 0; i < noSitStaffList.size(); i++) {
            Staff staff = noSitStaffList.get(i);
            int lowestWeight = -1;
            int lowestWeightTableId = -1;
            int lowestWeightCount = 0;
            //计算餐桌对员工权重，得到最小权重的餐桌
            for (int j = 0; j < tableArrayList.size(); j++) {
                Table table = tableArrayList.get(j);
                if (table.currentCount <= tableCount) {
                    int oldWeight = 0;
                    if (table.weightMap.containsKey(staff.status)) {
                        oldWeight = table.weightMap.get(staff.status) * fanxu(OLD_PRIORITY);
                    }
                    int sexWeight = 0;
                    if (table.weightMap.containsKey(staff.sex)) {
                        sexWeight = table.weightMap.get(staff.sex) * fanxu(SEX_PRIORITY);
                    }
                    int sectionWeight = 0;
                    if (table.weightMap.containsKey(staff.section)) {
                        sectionWeight = table.weightMap.get(staff.section) * fanxu(SECTION_PRIORITY);
                    }
                    int currentTable2StaffWeight = oldWeight + sexWeight + sectionWeight;
                    if (lowestWeight == -1) {
                        lowestWeight = currentTable2StaffWeight;
                        lowestWeightTableId = j;
                    }
                    if (currentTable2StaffWeight < lowestWeight) {
                        lowestWeightCount = 1;
                        lowestWeight = currentTable2StaffWeight;
                        lowestWeightTableId = j;
                    }
                    if (currentTable2StaffWeight == lowestWeight) {
                        lowestWeightCount = lowestWeightCount + 1;
                    }
                }
            }
            //如果是当前第一个未分配座位的员工，选择权重最小的餐桌（可能有多个，选第一个）；
            //如果非第一个非分配座位的员工，选中最小餐桌只有一个的话，直接选择，如果有多个，跳过该员工，等待下次轮询
            if (!staff.isSit && (lowestWeight != -1 && lowestWeightCount == 1 || (lowestWeightCount > 1 && i == 0))) {
                Table table = tableArrayList.get(lowestWeightTableId);
                if (table.weightMap.containsKey(staff.status)) {
                    table.weightMap.put(staff.status, table.weightMap.get(staff.status) + 1);
                } else {
                    table.weightMap.put(staff.status, 1);
                }
                if (table.weightMap.containsKey(staff.sex)) {
                    table.weightMap.put(staff.sex, table.weightMap.get(staff.sex) + 1);
                } else {
                    table.weightMap.put(staff.sex, 1);
                }
                if (table.weightMap.containsKey(staff.section)) {
                    table.weightMap.put(staff.section, table.weightMap.get(staff.section) + 1);
                } else {
                    table.weightMap.put(staff.section, 1);
                }
                table.staffArrayList.add(staff);
                table.currentCount = table.currentCount + 1;
                staff.isSit = true;
            }
        }
        if (noSitStaffList.size() == 0) {
//            Gson gson = new Gson();
//            String json = gson.toJson(tableArrayList);
//            Log.e("->MainActivity", "chooseTable@99 --> " + json);
            try {
                ExcelUtil.writeExcel("app/xlsx2xml/分桌后.xlsx", "Sheet1", tableArrayList, NUMBER_OF_PER_TABLE);
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
//            for (Table table : tableArrayList) {
//                Log.e("->MainActivity", "chooseTable@109 --> 餐桌" + table.tableId + "人员：" + table.staffArrayList.toString());
//            }
        } else {
            //递归
            chooseTable(noSitStaffList);
        }
    }

    private static ArrayList<Staff> getNoSitStaffList(ArrayList<Staff> staffArrayList){
        ArrayList<Staff> temp=new ArrayList<>();
        for (Staff staff : staffArrayList) {
            if (!staff.isSit) {
                temp.add(staff);
            }
        }
        return temp;
    }

    private static int fanxu(int i) {
        return 3 - (i + 1) + 1;
    }

}
