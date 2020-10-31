
		milestone_text_style = document.styles['Normal']
		milestone_text_font = milestone_text_style.font
		milestone_text_font.name = 'Arial'
		milestone_text_font.size = Pt(16)
		milestone_text_font.bold = True

# Assign the theme color based on the category from column A of the excel sheet.
		if category == 'MACHINES':
			milestone_text_font.color.rgb = RGBColor(0x3f, 0x2c, 0x36)
		if category == 'INNOVATIONS':
			milestone_text_font.color.rgb = RGBColor(0x42, 0x24, 0xE9)
		if category == 'PEOPLE':
			milestone_text_font.color.rgb = RGBColor(0x36, 0x24, 0x3f)
