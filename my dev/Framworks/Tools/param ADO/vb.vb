If is_Delete = False Then
    'insert
        colParams.Add addParam("projID", gVarchar, "100", en.p_projId)
        colParams.Add addParam("task", gVarchar, "100", en.p_task)
        colParams.Add addParam("assignee", gVarchar, "100", CStr(en.p_assignee))
        colParams.Add addParam("deliverable", gVarchar, "100", en.p_deliverable)
        colParams.Add addParam("description", gVarchar, "100", en.p_description)
        colParams.Add addParam("estStart", gVarchar, "100", en.p_estStart)
        colParams.Add addParam("estEnd", gVarchar, "100", en.p_estEnd)
        colParams.Add addParam("actStart", gVarchar, "100", en.p_actStart)
        colParams.Add addParam("actEnd", gVarchar, "100", en.p_actEnd)
        colParams.Add addParam("prio", gVarchar, "100", en.p_priority)
        colParams.Add addParam("statuss", gVarchar, "100", en.p_status)
        If is_Edit = True Then
            'update
            colParams.Add addParam("ids", gVarchar, "100", en.p_id)
        End If
    Else
	
	
	
"SELECT BT_Status.Proj_Status, BT_ProjDetails.Proj_Priority, BT_ProjDetails.Est_Start, BT_ProjDetails.Est_End, DateDiff("d", BT_ProjDetails.Est_Start, BT_ProjDetails.Est_End) AS [Est Days], BT_ProjDetails.Task, BT_Assignee.Assignee_name, BT_ProjDetails.Deliverable, BT_ProjDetails.Description, BT_ProjDetails.Act_Start, BT_ProjDetails.Act_End, DateDiff("d", BT_ProjDetails.Act_Start, BT_ProjDetails.Act_End) AS [act Days], BT_ProjDetails.proj_id FROM (BT_ProjDetails LEFT JOIN BT_Assignee ON BT_ProjDetails.Assignee_ID = BT_Assignee.id) LEFT JOIN BT_Status ON BT_ProjDetails.Status_id = BT_Status.id"
