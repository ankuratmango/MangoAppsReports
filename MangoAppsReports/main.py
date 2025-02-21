import sys
import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image
from db_helper import DatabaseConnection
from chartgenerator import ChartGenerator
from collections import Counter

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

    
    issuer_count = Counter(entry['message_by'] for entry in recognition_hash.values())
    chart_data_issuers = list(issuer_count.items())
    chart_data = [(v['username'], int(v['award_count'])) for v in awardees_hash.values()]

    xls_data['chart_data_issuers'] = chart_data_issuers
    xls_data['chart_data'] = chart_data
    
    output_path="static_chart.xlsx"
    generator = ChartGenerator(xls_data, output_path)
    generator.generate_excel_summary("Summary")

    headers = [
            "Award Name", "Category", "Message", "Given By", "Employee ID of Given By", "Approved By", 
            "Given To", "Employee Id of Given To", "Manager", "Manager Employee Id", "Given On (mm/dd/yyyy)", 
            "Gamification Points", "Reward Points", "Total Reward Points", "Departments", "Departments2"
    ]

    generator.generate_excel_data("Data", headers)
   
except Exception as ex:
    print(ex)
finally:
    db.disconnect()

