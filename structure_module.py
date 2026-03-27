# -*- coding: utf-8 -*-
"""
Created on Mon Oct 20 11:16:43 2025

@author: AAA
"""

from PyQt5.QtWidgets import QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QHBoxLayout, QFileDialog

class StructureModule(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("林分结构参数估计模块")
        self.resize(600, 400)

        layout = QVBoxLayout()

        # 输入TIF
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("输入TIF文件夹:"))
        self.input_folder = QLineEdit()
        row1.addWidget(self.input_folder)
        btn_in = QPushButton("浏览")
        row1.addWidget(btn_in)
        layout.addLayout(row1)
        btn_in.clicked.connect(self.browse_input)

        # 输出TIF
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("输出TIF路径:"))
        self.output_tif = QLineEdit()
        row2.addWidget(self.output_tif)
        btn_out = QPushButton("设置")
        row2.addWidget(btn_out)
        layout.addLayout(row2)
        btn_out.clicked.connect(self.set_output)

        # 日志和运行按钮
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(QLabel("运行日志:"))
        layout.addWidget(self.log)
        self.run_btn = QPushButton("开始估计")
        layout.addWidget(self.run_btn)
        self.run_btn.clicked.connect(self.run_processing)

        self.setLayout(layout)

    def browse_input(self):
        path = QFileDialog.getExistingDirectory(self, "选择TIF文件夹")
        if path:
            self.input_folder.setText(path)

    def set_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存输出TIF", "", "TIF文件 (*.tif)")
        if path:
            self.output_tif.setText(path)

    def log_msg(self, msg):
        self.log.append(msg)

    def run_processing(self):
        self.log.clear()
        self.log_msg("开始林分结构参数估计...")
        self.log_msg(f"输入文件夹: {self.input_folder.text()}")
        self.log_msg(f"输出TIF: {self.output_tif.text()}")
        self.log_msg("处理完成（此处可加入实际林分参数计算算法）")