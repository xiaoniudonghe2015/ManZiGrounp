package com.example.fenzhuo;

import java.util.ArrayList;
import java.util.HashMap;

public class Table {
    public String tableId;//餐桌号
    public ArrayList<Staff> staffArrayList = new ArrayList<>();//每桌员工信息
    public HashMap<String, Integer> weightMap = new HashMap<>();//记录各属性条件的权重
    public int currentCount;//当前状态餐桌的人数
}
