import xlwt


class FontColor(object):

    def __init__(self):
        super(FontColor, self).__init__()

        # -------------------------------------------------------
        # Styles for Excel sheet Row, Column, Text - color, Font
        # -------------------------------------------------------
        self.style0 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                  'font: name Arial, color black, bold on;')
        self.style1 = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;'
                                  'font: name Arial, color black, bold off;')
        self.style2 = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                                  'font: name Arial, color yellow, bold on;')
        self.style3 = xlwt.easyxf('font: name Arial, color red, bold on')
        self.style4 = xlwt.easyxf('pattern: pattern solid, fore_colour indigo;'
                                  'font: name Arial, color gold, bold on;'
                                  'align: vert centre, horiz centre;')
        self.style5 = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;'
                                  'font: name Arial, color brown, bold on;'
                                  'align: vert centre, horiz centre;')
        self.style6 = xlwt.easyxf('font: name Arial, color light_orange, bold on')
        self.style7 = xlwt.easyxf('font: name Arial, color orange, bold on')
        self.style8 = xlwt.easyxf('font: name Arial, color green, bold on')
        self.style9 = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;'
                                  'font: name Arial, color brown, bold on;')
        self.style10 = xlwt.easyxf('pattern: pattern solid, fore_colour brown;'
                                   'font: name Arial, color yellow, bold on;')
        self.style11 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.style12 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.style13 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.style14 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.style15 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                   'font: name Times New Roman, color-index black, bold on')
        self.style16 = xlwt.easyxf('font: name Arial, color black, bold off')
        self.style17 = xlwt.easyxf('font: name Arial, color light_orange')
        self.style18 = xlwt.easyxf('font: name Arial, color red')
        self.style19 = xlwt.easyxf('pattern: pattern solid, fore_colour gold;'
                                   'font: name Arial, color black, bold on;')
        self.style20 = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;'
                                   'font: name Arial, color dark_red_ega, bold off;')
        self.style21 = xlwt.easyxf('pattern: pattern solid, fore_colour periwinkle;'
                                   'font: name Arial, color black, bold on;')
        self.style22 = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow;'
                                   'font: name Arial, color black, bold on;')
        self.style23 = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                                   'font: name Arial, color light_orange, bold on, height 300; '
                                   'align:wrap on, vert centre, horiz centre;')
        self.style24 = xlwt.easyxf('font: name Arial, color green, bold on, height 400;'
                                   'align: vert centre, horiz centre;')
        self.style25 = xlwt.easyxf('font: name Arial, color red, bold on, height 400;'
                                   'align: vert centre, horiz centre;')
        self.style26 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on;'
                                   'align: vert centre, horiz centre;')

# -------
# colours
# -------
# aqua 0x31
# black 0x08
# blue 0x0C
# blue_gray 0x36
# bright_green 0x0B
# brown 0x3C
# coral 0x1D
# cyan_ega 0x0F
# dark_blue 0x12
# dark_blue_ega 0x12
# dark_green 0x3A
# dark_green_ega 0x11
# dark_purple 0x1C
# dark_red 0x10
# dark_red_ega 0x10
# dark_teal 0x38
# dark_yellow 0x13
# gold 0x33
# gray_ega 0x17
# gray25 0x16
# gray40 0x37
# gray50 0x17
# gray80 0x3F
# green 0x11
# ice_blue 0x1F
# indigo 0x3E
# ivory 0x1A
# lavender 0x2E
# light_blue 0x30
# light_green 0x2A
# light_orange 0x34
# light_turquoise 0x29
# light_yellow 0x2B
# lime 0x32
# magenta_ega 0x0E
# ocean_blue 0x1E
# olive_ega 0x13
# olive_green 0x3B
# orange 0x35
# pale_blue 0x2C
# periwinkle 0x18
# pink 0x0E
# plum 0x3D
# purple_ega 0x14
# red 0x0A
# rose 0x2D
# sea_green 0x39
# silver_ega 0x16
# sky_blue 0x28
# tan 0x2F
# teal 0x15
# teal_ega 0x15
# turquoise 0x0F
# violet 0x14
# white 0x09
# yellow 0x0D


# --------------
# Text alignment
# --------------
# xlwt.easyxf('align:vert top, horiz right')
# xlwt.easyxf('align:wrap on; font: bold on, color-index red')
# xlwt.easyxf('align: vert bottom, horiz left')
# xlwt.easyxf('align: vert centre, horiz centre')
