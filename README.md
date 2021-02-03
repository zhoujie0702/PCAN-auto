# PCAN-auto
因为公司没有这种支持自定义发送脚本的CAN工具,所幸自己会一点Python,所以工作之余写了这么一个脚本,只是为了增加工作效率.
当然,因为python初学,代码不仅乱还粗劣,要是有哪里可以继续改进的话还请各位大神多多指点

通过Python脚本,借用PCAN提供的动态库来实现自定义的CAN数据发送

PS1:
sheet创建:CAN数据的导入需要通过DATA.xlsx表格,表格中的每个sheet对应的是每个信号ID,理论上可以无限创建ID
msg.REPEAT_EVERY:该ID下每条报文的重复发送次数
msg.ID:该报文的ID
msg.CYCLE:该ID下报文的周期
msg.REPEAT_ALL:该ID下报文循环发送第一条至最后一条的次数
