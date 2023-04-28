import yaml
import datetime
import calendar
import chinese_calendar
mdict = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr",
         5: "May", 6: "Jun", 7: "Jul", 8: "Aug",
         9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}


class MailTable(object):
    records = None
    html = None
    table_head = None
    td_head = None

    date = None
    month = None
    year = None
    end_day = None
    workdays = None

    def __init__(self, send_date):
        # self.month = send_date.month - 1 if send_date.month - 1 != 0 else 12
        self.month = send_date.month
        # self.year = send_date.year if send_date.month - 1 != 0 else send_date.year - 1
        self.year = send_date.year
        self.end_day = calendar.monthrange(self.year, self.month)[1]
        self.workdays = chinese_calendar.get_workdays(start=datetime.date(self.year, self.month, 1),
                                                 end=datetime.date(self.year, self.month, self.end_day))
        self.records = self.get_leave_records_with_month(self.month)
        self.html = "cc: chenxiao@wutron.com;liangjl@wutron.com<br> <p>Hi all,</p><p>Attendance records here. </p>"
        self.table_head = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=1113 style='width:834.6pt;border-collapse:collapse'>"
        self.templete_tr = """
            <tr style='height:15.6pt'>
                <td width=50 nowrap style='width:37.8pt;border:solid black 1.0pt;border-bottom:none;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'>
                    <p class=MsoNormal align=center style='text-align:center'>
                        <b><span lang=ZH-CN style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>姓名</span></b>
                        <b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></b>
                    </p>
                </td>
                <td width=65 style='width:48.6pt;border-top:solid black 1.0pt;border-left:none;border-bottom:none;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'><p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>Virdis Wu<o:p></o:p></span></p></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=24 nowrap style='width:.25in;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=82 nowrap style='width:61.8pt;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=65 nowrap style='width:48.6pt;border:none;border-top:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
                <td width=106 nowrap style='width:79.8pt;border-top:solid black 1.0pt;border-left:none;border-bottom:none;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.6pt;border-image: initial'></td>
            </tr>
            <tr style='height:33.0pt'>
                <td nowrap rowspan=2 style='border:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:33.0pt;border-image: initial'><p class=MsoNormal align=center style='text-align:center'><b><span lang=ZH-CN style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>月份</span></b><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></b></p></td>
                <td nowrap rowspan=2 style='border:solid black 1.0pt;border-left:none;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:33.0pt'><p class=MsoNormal align=center style='text-align:center'><b><span lang=ZH-CN style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>时间</span></b><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></b></p></td>
                <td nowrap colspan=31 style='border:solid black 1.0pt;border-left:none;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:33.0pt'><p class=MsoNormal align=center style='text-align:center'><b><span lang=ZH-CN style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>出勤情况</span></b><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></b></p></td>
                <td width=82 rowspan=2 style='width:61.8pt;border:solid black 1.0pt;border-left:none;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:33.0pt'><p class=MsoNormal align=center style='text-align:center'><span class=font1><b><span lang=ZH-CN style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>本月缺勤</span></b></span><span class=font1><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>DAY/HOUR</span></b></span><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></b></p></td>
                <td width=65 rowspan=2 style='width:48.6pt;border:solid black 1.0pt;border-left:none;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:33.0pt'><p class=MsoNormal align=center style='text-align:center'><span class=font1><b><span lang=ZH-CN style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>本月加班</span></b></span><span class=font1><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>DAY/HOUR</span></b></span><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></b></p></td>
                <td width=106 rowspan=2 style='width:79.8pt;border:solid black 1.0pt;border-left:none;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:33.0pt;border-image: initial'><p class=MsoNormal align=center style='text-align:center'><b><span lang=ZH-CN style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>备注</span></b><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></b></p></td>
            </tr>
            <tr style='height:15.6pt'>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>1<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>2<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>3<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>4<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>5<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>6<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>7<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>8<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>9<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>10<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>11<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>12<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>13<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>14<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>15<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>16<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>17<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>18<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>19<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>20<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>21<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>22<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>23<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>24<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>25<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>26<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>27<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>28<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>29<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>30<o:p></o:p></span></b></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;background:#C5D9F1;padding:.75pt .75pt 0in .75pt;height:15.6pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:10.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>31<o:p></o:p></span></b></p></td>
            </tr>
            
            <tr style='height:.25in'>
                <td nowrap rowspan=2 style='border:solid black 1.0pt;border-top:none;padding:.75pt .75pt 0in .75pt;height:.25in;border-image: initial;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'><p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>{}<o:p></o:p></span></p></td>
                <td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:.25in;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'><p class=MsoNormal align=center style='text-align:center'>
                    <span lang=ZH-CN style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>上午</span>
                    <span style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></p>
                </td>
        """.format(mdict[self.month])
        self.html += self.table_head
        self.html += self.templete_tr

        self.add_record_html("AM")

        self.html += self.add_summary()

        self.add_record_html("PM")

    def add_record_html(self, when):
        for i in range(1, 32):
            if self.is_leave(i, "ALL") or self.is_leave(i, when):
                self.html += self.add_record(2)
            else:
                if i in [workday.day for workday in self.workdays]:
                    self.html += self.add_record(0)
                else:
                    self.html += self.add_record(1)

    def add_record(self, work_flag):
        """
        :param work_flag:
            0-> workday: ✓
            1-> holiday: (None)
            2-> leave:   x
        :return:
        """

        flag = None
        if work_flag == 0:
            flag = "✓"
        if work_flag == 1:
            flag = ""
        if work_flag == 2:
            flag = "x"
        td = "<td width=24 style='width:.25in;border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.15pt;border-image: initial;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'>{}</td>".format(flag)
        return td

    def add_summary(self, leave_day="", ot_day="", notes=""):
        """
        本月缺勤hour写入点
        本月加班hour写入点
        备注写入点
        """
        summary_day, notes = self.sum_days(self.records)
        context = """
        <td nowrap rowspan=2 style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:.25in;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'>
            <p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>{}<o:p></o:p></span></p>
        </td>
        <td width=65 rowspan=2 style='width:48.6pt;border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:.25in;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'>
            <p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>{}<o:p></o:p></span></p>
        </td>
        <td width=106 rowspan=2 style='width:79.8pt;border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:.25in;border-image: initial;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'>
            <p class=MsoNormal align=center style='text-align:center'><span style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>{}<o:p></o:p></span></p>
        </td>
    </tr>
    <tr style='height:15.15pt'><td nowrap style='border-top:none;border-left:none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;padding:.75pt .75pt 0in .75pt;height:15.15pt;border-image: initial;background-image:initial;background-position:initial;background-size: initial;background-repeat:initial;background-attachment:initial;background-origin: initial;background-clip: initial'><p class=MsoNormal align=center style='text-align:center'><span lang=ZH-CN style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'>下午</span><span style='font-size:9.0pt;font-family:"Microsoft YaHei",sans-serif;color:black'><o:p></o:p></span></p></td>
        """.format(summary_day, ot_day, notes)
        return context

    def get_html(self):
        self.html += '</tr></table>'
        self.html += '<p>BR/{0}</p>'.format("Virdis")
        return self.html

    def is_leave(self, date_time, when):
        for r in self.records:
            if date_time == r.get("date").day and when == r.get("when"):
                print("Find leave record:", r.get("date"), r.get("when"))
                return True
        return False

    def sum_days(self, target_record):
        notes = []
        day = 0.0
        for r in target_record:
            if self.month == r.get("date").month:
                if r.get("when") == "ALL":
                    day += 1
                    notes.append("{}请一天年假".format(datetime.datetime.strftime(r.get("date"), "%Y-%m-%d")))
                else:
                    day += 0.5
                    notes.append("{}请半天年假".format(datetime.datetime.strftime(r.get("date"), "%Y-%m-%d")))
        summary_day = "{}".format("0.0")
        return summary_day, "\n".join(notes)

    @staticmethod
    def get_leave_records_with_month(month):
        filename = "records.yaml"
        with open(filename, 'r', encoding='utf-8') as fr:
            data = fr.read()
        results = yaml.load(data, Loader=yaml.FullLoader)["leave"]
        for result in results:
            d = datetime.datetime.strptime(result.get("date"), "%Y-%m-%d")
            result["date"] = d
        records = []
        for r in results:
            if r.get("date").month == month:
                records.append(r)
        return records

    def get_subject(self):
        return "{}月考勤记录".format(self.month)




