from fpdf import FPDF
import csv

def write_cover_letter(cover_letter, skills):

    # open csv file and read input
    with open(skills) as skills_csv:

        reader = csv.reader(skills_csv)
        rownum = 0

        for row in reader:

            pdf = FPDF('P', 'mm', 'A4')  # portrait mode, mm , A4 size paper
            pdf.add_page()  # new blank page
            pdf.set_font('Arial', '', 12)  # font, Style (B,U,I) , fontsize in pt.

            #ignore the header row
            if rownum == 0:
                pass

            else:
                model_cover_letter = open(cover_letter, 'r')

                for line in model_cover_letter:

                    line = line.replace('#website', row[0])
                    line = line.replace('#inserttools', ','.join(row[1].split('#')))  # skills are seperated by '#' split and join them
                    line = line.replace('#toolproficient', row[2])
                    line = line.replace('#toolyr', row[3])
                    line = line.replace('#company', row[4])

                    pdf.write(6, line)

                pdf.output('Cover Letter - ' + row[4] + '.pdf', 'F')
                pdf.close()

            rownum = rownum + 1

if __name__ == "__main__":

    coverletter = 'cover_letter.txt'
    skillset = 'skills.csv'

    # just use the right file names or modify the ones provided
    write_cover_letter(coverletter, skillset)
