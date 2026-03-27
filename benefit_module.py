# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 16:34:30 2025

@author: AAA
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QTextEdit
import numpy as np

class BenefitModule(QWidget):
    def __init__(self, mode=None):
        super().__init__()
        self.setWindowTitle(f"效益分析 - {mode or '综合'}")
        self.resize(600, 400)
        layout = QVBoxLayout(self)
        self.text = QTextEdit()
        layout.addWidget(self.text)

        if mode == "wood":
            self.wood_benefit()
        elif mode == "eco":
            self.eco_benefit()
        elif mode == "econ":
            self.econ_benefit()

    def wood_benefit(self):
        v = np.array([30, 45, 60])  # 木材体积（m³/ha）
        price = 350  # 元/m³
        total = np.sum(v * price)
        self.text.setText(f"木材效益估算：\n总产量 {v.sum()} m³/ha\n总价值 {total:.2f} 元/ha")

    def eco_benefit(self):
        c_fix = 2.5  # tC/ha
        carbon_price = 300  # 元/tC
        value = c_fix * carbon_price
        self.text.setText(f"碳汇价值估算：\n固碳量 {c_fix} tC/ha\n价值 {value:.2f} 元/ha")

    def econ_benefit(self):
        gdp = 10000  # 假设地区GDP贡献
        ratio = 0.05
        benefit = gdp * ratio
        self.text.setText(f"经济效益估算：\n区域GDP: {gdp} 元\n森林经济贡献: {benefit:.2f} 元")
