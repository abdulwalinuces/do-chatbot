import pyodbc 
import re
import datetime
import sys, os


def get_connection():
    connection_string = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=.\SQLEXPRESS;DATABASE=EB_DOP;UID=root;PWD=*******"
    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()
    return conn, cursor

def search_query(query):
    """
    run the query on mssql database.
    """
    conn, cursor =  get_connection()
    cursor.execute(query)
    results = cursor.fetchall()
    # Close the connection
    conn.close()
    return results

def managedata(user_email):
  try:
    conn, cursor =  get_connection()
    user_sid = get_user_sid(user_email, cursor)
    # Get pending validations
    pending_list = []
    cursor.execute("""SELECT r.NetworkPath  
                      FROM EB_DOP_ResourceOwners o
                      JOIN EB_DOP_Resources r ON o.ResourceId_Id = r.Id
                      WHERE o.UserSId_Id = ? AND o.ValidationStatusId_id = ?""",  
                    (user_sid, get_status_id('Pending', cursor)))
    
    for row in cursor:
        pending_list.append(row.NetworkPath)

    # Get review history       
    history_list = []
    cursor.execute("""SELECT r.NetworkPath
                      FROM EB_DOP_Reviews s  
                      JOIN EB_DOP_Resources r ON s.ResourceId_Id = r.Id
                      WHERE s.Id IN (SELECT ReviewId_Id 
                                     FROM EB_DOP_ReviewResponses_DOV
                                     WHERE ReviewerSId = ? AND Value IS NULL)""",  
                    (user_sid))
                                    
    for row in cursor:
        history_list.append(row.NetworkPath)

    data = {
      "pending": pending_list,
      "history": history_list 
    }
    conn.close()
    return data
  
  except Exception as e:
    print(e)
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)

def get_invalid_resources():
    query= "SELECT  EB_DOP_ResourceOwners.Resourceid_id, EB_DOP_ResourceOwners.UserSId_id, EB_DOP_ResourceOwners.ValidationStatusId_id, EB_DOP_Resources.NetworkPath, EB_DOP_Resources.ParentResourceId, ResourceTypeId_id " 
    query +=  " FROM EB_DOP_ResourceOwners INNER JOIN EB_DOP_Resources ON EB_DOP_ResourceOwners.ResourceId_id = EB_DOP_Resources.id where EB_DOP_ResourceOwners.ValidationStatusId_id != 2"
    results = search_query(query)
    resources = {}
    for r in results:
        if r[1] not in resources:
            resources[r[1]] = [r[3]]
        else:
            resources[r[1]].append(r[3])
    return resources

def get_resource_id(resource_path):
    try:
        query= "SELECT id from EB_DOP_Resources where NetworkPath = '"+resource_path+"'"
        results = search_query(query)
        if results:
            return results[0][0]
        return None
    except Exception as ex:
        print(ex)
        return None
    
def get_resource_path(id_):
    try:
        query= "SELECT NetworkPath from EB_DOP_Resources where id = '"+id_+"'"
        results = search_query(query)
        if results:
            return results[0][0]
        return None
    except Exception as ex:
        print(ex)
        return None

def get_unaccepted_resources():
    """
    Reads resources by owner from mssql database, which are not accepted yet.
    """
    try:
        query= "SELECT  EB_DOP_ResourceOwners.Resourceid_id, EB_DOP_ResourceOwners.UserSId_id, EB_DOP_ResourceOwners.ValidationStatusId_id, EB_DOP_Resources.NetworkPath, EB_DOP_Resources.ParentResourceId, ResourceTypeId_id " 
        query +=  " FROM EB_DOP_ResourceOwners INNER JOIN EB_DOP_Resources ON EB_DOP_ResourceOwners.ResourceId_id = EB_DOP_Resources.id where EB_DOP_ResourceOwners.ValidationStatusId_id = 1"
        results = search_query(query)
        resources = {}
        for r in results:
            if r[1] not in resources:
                resources[r[1]] = [r[3]]
            else:
                resources[r[1]].append(r[3])
        return resources
    except Exception as ex:
       print(ex)

def get_owner_details(Sid):
    query =  "Select * from EB_DOP_UserPrincipals where UserSId = '"+Sid+"'"
    results = search_query(query)
    if results:
        return results[0]
    return None

def get_owner_details_by_email(email_id):
    query =  "Select * from EB_DOP_UserPrincipals where EmailAddress = '"+email_id+"'"
    results = search_query(query)
    if results:
        return results[0]
    return None

def check_owners_exists(owner):
    """ Check the owner confirmed or suggested by owner in the email response"""
    try:
        if re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', owner['email']) is not None:
            # if email id is available for owner the check email id in database
            query =  "Select * from EB_DOP_UserPrincipals where EmailAddress = '"+owner['email']+"'"
            results = search_query(query)
            # if the owner name and 
            if results:
                owner['status'] = 'valid'
                return owner
        else:
            query = "SELECT * FROM EB_DOP_UserPrincipals WHERE CHARINDEX('"+owner['name']+"', DisplayName) > 0;"
            results = search_query(query)
            owner['status'] = 'reconfirm'
            if results:
                owner['names'] = []
                for result in results:
                    owner['names'].append((result[1], result[3], result[4]))
            else:
                owner['names'] = []
                owner['names'].append((owner['name'], '', ''))
                owner['status'] = 'reconfirm'
            return owner
    except Exception as ex:
        print(ex)
        import os, sys
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)


def update_resource_accept(user_email, folder_path):
  try:
    conn, cursor =  get_connection()

    # Get resource and owner
    cursor.execute("SELECT Id FROM EB_DOP_Resources WHERE NetworkPath = ?", folder_path)
    resource_id = cursor.fetchone()[0]
    
    cursor.execute("SELECT Id FROM EB_DOP_ResourceOwners WHERE ResourceId_Id = ? AND UserSId_Id = ?",  
                   (resource_id, get_user_sid(user_email, cursor)))
    owner_id = cursor.fetchone()[0]

    # Update owner validation
    cursor.execute("UPDATE EB_DOP_ResourceOwners SET ValidationStatusId_Id = ?, ConfirmTime = ? WHERE Id = ?",  
                   (get_validation_status_id('Accepted', cursor), datetime.datetime.now(), owner_id))

    # Update review response
    cursor.execute("UPDATE EB_DOP_ReviewResponses_DOV SET Value = ?, CompletionDate = ? WHERE ReviewId_Id = ? AND Value IS NULL",
                  ('Accepted', datetime.datetime.now(), get_review_id(resource_id))) 

    # Commit changes
    conn.commit()
    conn.close()
    return "Updated successfully"

  except Exception as e:
    print(e)
    return "Error updating acceptance"
  
def get_user_sid(email, cursor):
  sql = "SELECT UserSId FROM EB_DOP_UserPrincipals WHERE EmailAddress = ?"
  cursor.execute(sql, (email,))
  row = cursor.fetchone()
  if row:
    return row[0]
  return None

def get_validation_status_id(status, cursor):
  # Query to get validation status
  sql = "SELECT Id FROM EB_DOP_Type_ValidationStatus WHERE validationtype = ?"
  cursor.execute(sql, (status))
  row = cursor.fetchone()
  if row:
    print('returning -- ', row[0])
    return row[0]
  return None
 
def get_review_id(resource_id):
  conn, cursor =  get_connection()
  sql = "SELECT Id FROM EB_DOP_Reviews WHERE ResourceId_Id = ? AND ResponseDate IS NULL"
  cursor.execute(sql, (resource_id,))
  row = cursor.fetchone()
  conn.close()
  if row:
    return row[0]
  return None

def update_resource_reject(user_email, resource_path, message, suggested_owner, suggested_owner_details):
  conn, cursor =  get_connection()
  try:
    # Get resource and owner
    resource_id = get_resource_id(resource_path)
    suggestted_Sid = get_user_sid(user_email, cursor)
    owner_id = get_owner_id(resource_id, suggestted_Sid)
    # Update owner validation status
    reviewer_sid = get_user_sid(suggested_owner, cursor)
    cursor.execute("UPDATE EB_DOP_ResourceOwners SET ValidationStatusId_Id = ?, ConfirmTime = ?, UserSid_id = ? WHERE Id = ?",
                   (get_validation_status_id('Rejected', cursor), datetime.datetime.now(), reviewer_sid,  owner_id))

    # Update review response 
    review_id = get_review_id(resource_id)
    cursor.execute("UPDATE EB_DOP_ReviewResponses_DOV SET Value = ?, CompletionDate = ? WHERE ReviewId_Id = ? AND ReviewerSId = ?",
                   ('Rejected', datetime.datetime.now(), review_id, get_user_sid(user_email, cursor)))

    # Add new suggested owner 
    reviewer_sid = get_user_sid(suggested_owner, cursor)  
    cursor.execute("INSERT INTO EB_DOP_ReviewResponses_DOV (ReviewId_Id, ReviewerSId, CreatedDate) VALUES (?, ?, ?)",
                (review_id, reviewer_sid, datetime.datetime.now()))

    #Save message  
    cursor.execute("INSERT INTO EB_DOP_ResponseDetails_DOV (ReviewId_Id, ResponseId_Id, ValueType, Value) VALUES (?, ?, 'Message', ?)",
                  (review_id, get_last_response_id(review_id, conn, cursor), message))

    # Save suggested owner details  
    details = suggested_owner_details[0]+' - '+suggested_owner_details[1]+' - '+suggested_owner_details[2]
    cursor.execute("INSERT INTO EB_DOP_ResponseDetails_DOV (ReviewId_Id, ResponseId_Id, ValueType, Value) VALUES (?, ?, 'Message', ?)",
                   
                  (review_id, get_last_response_id(review_id, conn, cursor), details))

    # Save suggested owners
    reviewer_sid = get_user_sid(suggested_owner, cursor)
    cursor.execute("INSERT INTO EB_DOP_ResponseDetails_DOV (ReviewId_Id, ResponseId_Id, ValueType, Value) VALUES (?, ?, 'SuggestedOwner', ?)",
                (review_id, get_last_response_id(review_id, conn, cursor), reviewer_sid))
                

    # Commit changes
    conn.commit()
    conn.close()
    return "Updated successfully"
       
  except Exception as e:
    conn.close()
    print(e)
    import os, sys
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)
  

def get_owner_id(resource_id, user_sid):
    conn, cursor =  get_connection()
    sql = "SELECT Id FROM EB_DOP_ResourceOwners WHERE ResourceId_Id = ? AND UserSId_Id = ?"
    cursor.execute(sql, (resource_id, user_sid))
    row = cursor.fetchone()
    conn.close()
    if row:
        return row[0]
    return None

def get_status_id(status_name, cursor):
  sql = "SELECT Id FROM EB_DOP_Type_ValidationStatus WHERE validationtype = ?"
  cursor.execute(sql, (status_name,))
  row = cursor.fetchone()
  if row:
    return row[0]
  return None

def get_last_response_id(review_id, conn, cursor):
    sql = "SELECT Id FROM EB_DOP_ReviewResponses_DOV WHERE ReviewId_Id = ?"
    cursor.execute(sql, (review_id))
    row = cursor.fetchone()
    if row:
        return row[0]
    return None 

def get_suggested_owner(name, department, title):
    try:
        query = "SELECT EmailAddress FROM EB_DOP_UserPrincipals WHERE CHARINDEX('"+name+"', DisplayName) > 0 AND CHARINDEX('"+department+"', DepartmentName) > 0 AND CHARINDEX('"+title+"', TitleName) > 0"
        conn, cursor =  get_connection()
        cursor.execute(query)
        row = cursor.fetchone()
        conn.close()
        if row:
            return row[0]
        return None 
    except Exception as ex:
       print(ex)
