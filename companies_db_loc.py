import mysql.connector
import xlwt
from xlwt import Workbook
import re, os, glob, sys

#Connecting to factset
connection = mysql.connector.connect(
        user='actelitron',
        password='actelitron123',
        host='backenddevmysql.c3xnrjje2vet.eu-west-2.rds.amazonaws.com',
        database='factset',
        port=3306)

if connection:
    print("Success connecting to the database")
else:
    print("Failure of connecting to the database")

mycursor = connection.cursor()

#Query which finds all active main companies
mycursor.execute("""SELECT distinct cov.proper_name, bbg.bbg_ticker, struct.l1_name,
                struct.l2_name, struct.l3_name FROM sym_v1_sym_bbg bbg,
                rbics_v1_rbics_entity_focus foc, sym_v1_sym_coverage cov,
                rbics_v1_rbics_sec_entity sec, rbics_v1_rbics_structure struct
                where sec.fsym_id = cov.fsym_id
                and sec.factset_entity_id = foc.factset_entity_id
                and cov.fsym_id = cov.fsym_primary_equity_id
                and cov.active_flag = '1'
                and struct.l6_id = foc.l6_id
                and bbg.fsym_id = cov.fsym_primary_listing_id
                """)

comp_tick_lev = mycursor.fetchall()

# Creating a dictionary of companies
index = {}
for c in comp_tick_lev:
    tick = c[1]
    l1 = c[2]
    l2 = c[3]
    l3 = c[4]
    c = c[0]
    c = c.lower()
    c_tokens = c.split()
    if (len(c_tokens) > 0):
        if c_tokens[0] not in index:
            index[c_tokens[0]] = [c_tokens + [tick, l1, l2, l3]]
        else:
            if c_tokens not in index[c_tokens[0]]:
                index[c_tokens[0]].append(c_tokens + [tick, l1, l2, l3])
for k in index:
    index[k].sort(key=len, reverse=True)
    # print (index[k])


# Creating an excel workbook
wb = Workbook()
current_directory = os.getcwd()
report_folder = os.path.join(current_directory, 'REPORTS_nostemming')
k = 0

#Searching for the companies in reports using the index dictionary
for foldername in os.listdir(report_folder):
    print ("------------" + foldername + "----------------")
    k += 1
    sh = wb.add_sheet(foldername[:31])
    folder = os.path.join(report_folder, foldername)
    set_found = []
    tick_lev_found = []
    for file_name in glob.glob(os.path.join(folder, "*.txt")):
        with open(file_name, 'r') as myfile:
            data=myfile.read().lower().split()
            for i, token in enumerate(data):
                if (token != "morgan" and token != "stanley"):
                    if token in index:
                        for found in index[token]:
                            n_tokens = len(found[:-4])
                            if found[:-4] == data[i:i + n_tokens]:
                                if ((' '.join(found[:-4])) not in set_found):
                                    set_found.append(' '.join(found[:-4]))
                                    tick_lev_found.append(found[-4:])
    print ("Done!")
    sh.write(0, 0, "Main ID")
    sh.write(0, 1, "Company")
    sh.write(0, 2, "Ticker")
    sh.write(0, 3, "Level 1")
    sh.write(0, 4, "Level 2")
    sh.write(0, 5, "Level 3")
    sh.write(0, 6, "Related to ID")
    sh.write(0, 7, "Rel type")

    count = 1
    #Printing the main companies found in reports by theme
    for a in range(len(set_found)):
        sh.write(count, 0, a+1)
        sh.write(count, 1, set_found[a])
        sh.write(count, 2, tick_lev_found[a][0])
        sh.write(count, 3, tick_lev_found[a][1])
        sh.write(count, 4, tick_lev_found[a][2])
        sh.write(count, 5, tick_lev_found[a][3])
        sh.write(count, 6, "main")
        sh.write(count, 7, "main")
        count +=1

        #Looking for all related companies to one specific main company
        mycursor.execute("""
            SELECT distinct cov.proper_name, bbg.bbg_ticker, struct.l1_name, struct.l2_name, struct.l3_name, rel.rel_type
            FROM factset.ent_v1_ent_scr_relationships rel, factset.sym_v1_sym_coverage cov, factset.ent_v1_ent_scr_sec_entity sec,
            sym_v1_sym_bbg bbg,
            rbics_v1_rbics_entity_focus foc,
            rbics_v1_rbics_sec_entity secrb,
            rbics_v1_rbics_structure struct,
            factset.sym_v1_sym_coverage cov_main,
            factset.ent_v1_ent_scr_sec_entity sec_main
            WHERE
            rel.target_factset_entity_id = sec.factset_entity_id
            and cov.fsym_id = cov.fsym_primary_equity_id
            and sec.fsym_id = cov.fsym_id
            and cov.active_flag = '1'
            and secrb.fsym_id = cov.fsym_id
            and secrb.factset_entity_id = foc.factset_entity_id
            and struct.l6_id = foc.l6_id
            and bbg.fsym_id = cov.fsym_primary_listing_id
            and rel.source_factset_entity_id = sec_main.factset_entity_id
            and cov_main.fsym_id = cov_main.fsym_primary_equity_id
            and sec_main.fsym_id = cov_main.fsym_id
            and cov_main.active_flag = '1'
            and LOWER(cov_main.proper_name) = \"""" + set_found[a] +"\";")
        related = mycursor.fetchall()
        
        #Printing all the related companies to an excel file
        for rel_comp in related:
            sh.write(count, 0, "related")
            sh.write(count, 1, rel_comp[0])
            sh.write(count, 2, rel_comp[1])
            sh.write(count, 3, rel_comp[2])
            sh.write(count, 4, rel_comp[3])
            sh.write(count, 5, rel_comp[4])
            sh.write(count, 6, a+1)
            sh.write(count, 7, rel_comp[5])
            count += 1

wb.save('companies_by_themes_lowercase_levels_related.xls')
mycursor.close()
connection.close()
