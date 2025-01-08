import sys
import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image
from db_helper import DatabaseConnection

db = DatabaseConnection(host="localhost", user="root", password="root", database="mangoapps_dev")

def parse_recognition_name(recog_name):
    try:
        data = json.loads(recog_name)
        if "data" in data:
            return data["data"].get("recognition_name", ""), data["data"].get("category_name", "")
    except (json.JSONDecodeError, TypeError):
        return recog_name, None

def get_recognition_data(db, parse_recognition_name):
    query_recog = "select fid, Post, Post1, Given_By_Id, Given_By, given_by_emp_id, Given_To_Team, Given_To, given_to_ids, Given_To_EmpId, Given_On,Points, Reward_Points, Recog_Name, Recog_Category_Name, Total_Reward_Points from (select f.id as fid, substring(f.message, 1, 5000) as Post, substring(fp.title, 1, 5000) as Post1, u.id as Given_By_Id, u.name as Given_By, u.emp_id as given_by_emp_id, ' ' as Given_To_Team, GROUP_CONCAT(uf.name separator ', ') as Given_To, GROUP_CONCAT(uf.id separator ',') as given_to_ids, GROUP_CONCAT(uf.emp_id separator ', ') as Given_To_EmpId, f.created_at as Given_On, aw.points as Points, aw.reward_points as Reward_Points, fp.label_1 as Recog_Name, fp.label_2 as Recog_Category_Name, aw.total_reward_points as Total_Reward_Points from feeds f, feed_properties fp, users u, follow_list fl, users uf, awards aw where f.domain_id = 1 and f.created_by = u.id and f.id = fp.feed_id and f.id=aw.feed_id and u.domain_id = 1 and f .feed_type = 'ALL' and f.category = 'E' and f.id = fl.feed_id and f.conversation_id is NULL and fl .user_id = uf.id  and f.created_at >= '2024-11-01 07:00:00' and f.created_at <= '2025-12-01 07:59:59.999999' and f.is_deleted = false group by f.id UNION select f.id as fid, substring(f.message, 1, 5000) as Post, substring(fp.title, 1, 5000) as Post1, u.id as Given_By_Id, u.name as Given_By, u.emp_id as given_by_emp_id, c.name as Given_To_Team, GROUP_CONCAT(uf.name separator ', ') as Given_To, GROUP_CONCAT(uf.id separator ',') as given_to_ids, GROUP_CONCAT(uf.emp_id separator ', ') as Given_To_EmpId, f.created_at as Given_On, aw.points as Points, aw.reward_points as Reward_Points, fp.label_1 as Recog_Name, fp.label_2 as Recog_Category_Name, aw.total_reward_points as Total_Reward_Points from feeds f, feed_properties fp, users u, follow_list fl, conversations c, users uf, awards aw where f.domain_id = 1 and f.created_by = u.id and f.id = fp.feed_id and f.id=aw.feed_id and u .domain_id = 1 and f.feed_type = 'GRP' and f.category = 'E' and f.id = fl.feed_id and f .conversation_id is NOT NULL and f.conversation_id = c.id and fl.user_id = uf.id and f.created_at >= '2024-11-01 07:00:00' and f.created_at <= '2025-12-01 07:59:59.999999' and f.is_deleted = false group by f.id UNION select f.id as fid, substring(f.message, 1, 5000) as Post, substring(fp.title, 1, 5000) as Post1, u.id as Given_By_Id, u.name as Given_By, u.emp_id as given_by_emp_id, c.name as Given_To_Team, ' ' as Given_To, '' as given_to_ids, '' as Given_To_EmpId, f.created_at as Given_On, aw.points as Points, aw.reward_points as Reward_Points, fp.label_1 as Recog_Name, fp.label_2 as Recog_Category_Name, aw.total_reward_points as Total_Reward_Points from feeds f, feed_properties fp, users u, conversations c, awards aw where f.domain_id = 1 and f.created_by = u.id and f.id = fp.feed_id and f.id=aw.feed_id and u.domain_id = 1 and f.feed_type = 'GRP' and f.category = 'E' and f .conversation_id is NOT NULL and f.conversation_id = c.id and c.domain_id = 1  and f.created_at >= '2024-11-01 07:00:00' and f.created_at <= '2025-12-01 07:59:59.999999' and f.is_deleted = false group by f.id UNION SELECT f.id AS fid, Substring(aw.message, 1, 5000) AS Post, ''  AS Post1, u.id AS Given_By_Id, u.name AS Given_By, u.emp_id AS given_by_emp_id, ' ' AS Given_To_Team, Group_concat(uf.name SEPARATOR ', ') AS Given_To, Group_concat(uf.id SEPARATOR ',') AS given_to_ids, Group_concat(uf.emp_id SEPARATOR ', ') AS Given_To_EmpId, f.created_at AS Given_On, aw.points  AS Points, aw.reward_points AS Reward_Points, fa.action_data collate utf8_general_ci AS Recog_Name, '' AS Recog_Category_Name, aw.total_reward_points AS Total_Reward_Points FROM feeds f, users u, users uf, awards aw, awards_users au, feed_actions fa WHERE aw.domain_id = 1 AND f.feed_type = 'LST' AND f.id = aw.feed_id AND au.award_id = aw.id AND uf.id = au.user_id AND u.id = aw.from_user_id AND fa.feed_id = f.id AND u.domain_id = 1 AND f.sub_category = 'E' AND f.conversation_id IS NULL AND f.created_at >= '2024-11-01 07:00:00' AND f.created_at <= '2025-12-01 07:59:59.999999' AND f.is_deleted = false GROUP  BY aw.id UNION SELECT aw.id AS fid, '' AS Post, ''  AS Post1, u.id AS Given_By_Id, u.name AS Given_By, u.emp_id AS given_by_emp_id, ' ' AS Given_To_Team, Group_concat(uf.name SEPARATOR ', ') AS Given_To, Group_concat(uf.id SEPARATOR ',') AS given_to_ids, Group_concat(uf.emp_id SEPARATOR ', ') AS Given_To_EmpId, aw.created_at AS Given_On, ''  AS Points, aw.reward_points AS Reward_Points, '' AS Recog_Name, '' AS Recog_Category_Name, aw.total_reward_points AS Total_Reward_Points FROM awards aw, users u, users uf, awards_users au WHERE aw.domain_id = 1 AND aw.channel = 'A' AND aw.id = au.award_id AND au.user_id = uf.id AND aw.from_user_id = u.id AND u.domain_id = 1 AND aw.created_at >= '2024-11-01 07:00:00' AND aw.created_at <= '2025-12-01 07:59:59.999999' GROUP  BY aw.id ) t group by fid;"
    recognition_data = db.fetch_all(query_recog)
    print(recognition_data)
    query_award = "select userid, username, sum(AwardCount) as AwardCount from (select uf.id as userid, uf.name as username,count(uf.id) as AwardCount from feeds f, users u, follow_list fl, users uf where f.domain_id = 1 and f.created_by = u.id and u.domain_id = 1 and f.feed_type = 'ALL' and f.category = 'E' and f.id = fl.feed_id and f.conversation_id is NULL and fl.user_id = uf.id and f.created_at >='2024-11-01 07:00:00' and f.created_at <='2025-12-31 07:59:59.999999' and f.is_deleted = false group by uf.id UNION select uf.id as userid, uf.name as username,count(uf.id) as AwardCount from feeds f, users u, follow_list fl, conversations c, users uf where f.domain_id = 1 and f.created_by = u.id and u.domain_id = 1 and f.feed_type = 'GRP' and f.category = 'E' and f.id = fl.feed_id and f.conversation_id is NOT NULL and f.conversation_id = c.id and fl.user_id = uf.id  and f.created_at >= '2024-11-01 07:00:00' and f.created_at <= '2025-12-01 07:59:59.999999' and f.is_deleted = false group by uf.id UNION SELECT uf.id AS userid, uf.NAME AS username, Count(uf.id) AS AwardCount FROM feeds f, users u, users uf, awards aw, awards_users au, feed_actions fa WHERE aw.domain_id = 1 AND f.feed_type = 'LST' AND f.id = aw.feed_id AND au.award_id = aw.id AND uf.id = au.user_id AND u.id = aw.from_user_id AND fa.feed_id = f.id AND u.domain_id = 1 AND f.sub_category = 'E' AND f.conversation_id IS NULL AND f.created_at >= '2024-11-01 07:00:00' AND f.created_at <= '2025-12-01 07:59:59.999999' AND f.is_deleted = false GROUP BY au.user_id UNION SELECT uf.id AS userid, uf.NAME AS username, Count(uf.id) AS AwardCount FROM awards aw, users u, users uf, awards_users au WHERE aw.domain_id = 1 AND aw.channel = 'A' AND aw.id = au.award_id AND au.user_id = uf.id AND aw.from_user_id = u.id AND u.domain_id = 1 AND aw.created_at >= '2024-11-01 07:00:00' AND aw.created_at <= '2025-12-01 07:59:59.999999' GROUP BY au.user_id) t group by userid;"
    award_data = db.fetch_all(query_award)
    print(award_data)

    recognition_hash = {}
    
    for _, c in recognition_data.iterrows():
        recognition_name, recognition_category = parse_recognition_name(c['Recog_Name'])
        recognition_hash[c['fid']] = {
            'message': c['Post1'] + (f" - {c['Post']}" if pd.notna(c['Post']) and not pd.isna(c['Recog_Name']) else ""),
            'message_by_id': c['Given_By_Id'],
            'given_by_emp_id': c['given_by_emp_id'],
            'message_by': c['Given_By'],
            'team_name': c['Given_To_Team'],
            'message_to': c['Given_To'],
            'message_to_emp_id': c['Given_To_EmpId'],
            'message_given_on': c['Given_On'],
            'award_points': c['Points'],
            'award_reward_points': c['Reward_Points'],
            'award_recognition_name': recognition_name,
            'award_recognition_category': recognition_category or c['Recog_Category_Name'],
            'award_total_reward_points': c['Total_Reward_Points'],
            'receiver_ids': c['given_to_ids']
        }

    awardees_hash = {
        c['userid']: {
            "username": c['username'],
            "award_count": c['AwardCount']
        } for _, c in award_data.iterrows()
    }

    approver_query = f"""
    SELECT feeds.id, awards.approved_by, users.name AS approver_name
    FROM awards
    JOIN feeds ON feeds.id = awards.feed_id
    JOIN users ON users.id = awards.approved_by
    WHERE feeds.id IN ({', '.join(map(str, recognition_hash.keys()))}) AND awards.approved_by IS NOT NULL
    """
    approver_data = db.fetch_all(approver_query)

    approver_hash = {
        c['id']: {
            "approved_by": c['approved_by'],
            "approver_name": c['approver_name']
        } for _, c in approver_data.iterrows()
    }
    return recognition_hash, awardees_hash, approver_hash

def generate_multi_chart_reports_v2(xlsx_package, xls_data_recognized, xls_data_issuers):
    # Select the active worksheet
    sheet = xlsx_package.active
    sheet.title = "Recognitions Report"

    # Add header for the report
    sheet.append(['RECOGNITIONS REPORT'])
    sheet.append(['Dec 01, 2023 - Nov 30, 2024'])
    sheet.append([])  # Empty row for spacing

    # Create "Most Recognized Recipients" chart
    sheet.append(['Most Recognized Recipients'])
    sheet.append(['Name', 'Count'])
    for row in xls_data_recognized:
        sheet.append(row)

    chart1 = BarChart()
    chart1.title = "Most Recognized Recipients"
    chart1.x_axis.title = 'Name'
    chart1.y_axis.title = 'Count'

    data1 = Reference(sheet, min_col=2, min_row=5, max_row=4 + len(xls_data_recognized), max_col=2)
    categories1 = Reference(sheet, min_col=1, min_row=5, max_row=4 + len(xls_data_recognized))
    chart1.add_data(data1, titles_from_data=False)
    chart1.set_categories(categories1)

    sheet.add_chart(chart1, "E5")

    # Create "Top Issuing Users" chart
    start_row = 6 + len(xls_data_recognized)
    sheet.append([])
    sheet.append(['Top Issuing Users'])
    sheet.append(['Name', 'Count'])
    for row in xls_data_issuers:
        sheet.append(row)

    chart2 = BarChart()
    chart2.title = "Top Issuing Users"
    chart2.x_axis.title = 'Name'
    chart2.y_axis.title = 'Count'

    data2 = Reference(sheet, min_col=2, min_row=start_row + 2, max_row=start_row + 1 + len(xls_data_issuers), max_col=2)
    categories2 = Reference(sheet, min_col=1, min_row=start_row + 2, max_row=start_row + 1 + len(xls_data_issuers))
    chart2.add_data(data2, titles_from_data=False)
    chart2.set_categories(categories2)

    sheet.add_chart(chart2, f"E{start_row + 2}")


def generate_excel_report(xls_data):
    xlsx_package = Workbook()
    generate_multi_chart_reports_v2(xlsx_package, xls_data)
    xlsx_package.save("multi_chart_report.xlsx")

try:
    db.connect()
    recognition_hash, awardees_hash, approver_hash = get_recognition_data(db, parse_recognition_name)
    print(recognition_hash)
    print(awardees_hash)
    print(approver_hash)
    
    xls_data = {}
    xls_data['first_header'] = "RECOGNITIONS REPORT"
    xls_data['second_header'] = "Dec 01, 2023 - Nov 30, 2024"
    xls_data['header_font_size'] = 11

    top_award_givers_count_hash = {}
    top_award_givers_name_hash = {}

    for key, value in recognition_hash.items():
        recognition_by_name_key = value['message_by_id']  # Avoid using username as key, use userid instead

        if recognition_by_name_key not in top_award_givers_count_hash:
            top_award_givers_count_hash[recognition_by_name_key] = 0

        top_award_givers_count_hash[recognition_by_name_key] += 1
        top_award_givers_name_hash[recognition_by_name_key] = value['message_by']


    xls_data = [
        top_award_givers_count_hash,
        top_award_givers_name_hash
    ]

    # generate_excel_report(xls_data)

    #---------------------------------

    # Example data
    xls_data_recognized = [
        ['Namrata Puranik Puranik', 51],
        ['Gauri Puranik', 16],
        ['Aalkhimovich aalk', 16],
        ['Alumni User', 10],
        ['Ankur Tripathi', 7]
    ]

    xls_data_issuers = [
        ['Gauri Puranik', 84],
        ['Namrata Puranik Puranik', 21]
    ]

    # Generate the report
    xlsx_package = Workbook()
    generate_multi_chart_reports_v2(xlsx_package, xls_data_recognized, xls_data_issuers)
    xlsx_package.save("multi_chart_report.xlsx")
except Exception as ex:
    print(ex)
finally:
    db.disconnect()

