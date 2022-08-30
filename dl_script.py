import os
import sys
import time
import argparse
from datetime import datetime

try:
    import DaVinciResolveScript as dvr_script
except ImportError:
    # sys.path.append("%PROGRAMDATA%/Blackmagic Design/DaVinci Resolve/Support/Developer/Scripting/Modules")
    sys.path.append("C:/ProgramData/Blackmagic Design/DaVinci Resolve/Support/Developer/Scripting/Modules")
    import DaVinciResolveScript as dvr_script


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("database_type")
    parser.add_argument("database_name")
    parser.add_argument("project_name")
    # parser.add_argument("publish_path")
    # parser.add_argument("project_path")
    parser.add_argument("output_path")
    parser.add_argument("--folders", default="")
    parser.add_argument("--timeline", default="")
    # parser.add_argument("--format", default="")
    # parser.add_argument("--codec", default="")
    parser.add_argument("--render_preset", default="")
    parser.add_argument("--database_ip", default="")
    args = parser.parse_args()

    database_type = args.database_type
    database_name = args.database_name
    database_ip = args.database_ip

    project_name = args.project_name
    # publish_path = args.publish_path
    # project_path = args.project_path
    output_path = args.output_path
    folders = args.folders
    timeline_name = args.timeline
    # format_ = args.format
    # codec = args.codec
    render_preset = args.render_preset

    formatted_output_path = datetime.now().strftime(output_path)

    resolve = _connect_to_resolve()

    # If we want to open the database and render directly from that project
    if database_type and database_name:
        _load_database(resolve, database_type, database_name, database_ip)

    project = _load_project(resolve, project_name, folders)

    if timeline_name:
        _set_timeline(project, timeline_name)

    jobId = _setup_render_job(project, formatted_output_path, render_preset)
    _start_render(project, jobId)

def _load_database(resolve, database_type, database_name, database_ip="127.0.0.1"):
    print("Loading database:", database_name)
    project_manager = resolve.GetProjectManager()
    
    target_database = {
        "DbType": database_type, 
        "DbName": database_name, 
        "IpAddress": database_ip }

    # check if current database is the same as the target database
    current_database = project_manager.GetCurrentDatabase()
    
    if target_database != current_database:
        assert project_manager.SetCurrentDatabase(target_database), "Cannot load database."

def _connect_to_resolve():
    resolve = None
    i = 0
    while i < 5:
        resolve = dvr_script.scriptapp("Resolve")
        if resolve is not None:
            break
        print("Waiting for Resolve to start...")
        time.sleep(5)
        i += 1
    if resolve is None:
        raise RuntimeError("Could not connect to DaVinci Resolve. There may be a problem starting it, or you may be using the free version.")
    print("wait a bit more for resolve to become responsible")
    time.sleep(5)
    return resolve

def _load_project(resolve, project_name, folders):
    project_manager = resolve.GetProjectManager()
    time.sleep(1)

    if folders:
        folders = folders.replace("\\", "/")
        for folder in folders.split("/"):
            print("Opening folder:", folder)
            assert project_manager.OpenFolder(folder), "Cannot open folder."

    assert project_manager.LoadProject(project_name), "Cannot load project. Loading Failed."
    time.sleep(1)

    project = project_manager.GetCurrentProject()
    time.sleep(1)
    
    assert project.GetName() == project_name, "Cannot load project. Name Mismatch"
    
    return project

def _load_project_by_path(resolve, project_path, project_name):
    project_manager = resolve.GetProjectManager()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    assert project_manager.GotoRootFolder(), "Cannot goto root folder."
    assert project_manager.CreateFolder(timestamp), "Cannot create folder."
    assert project_manager.OpenFolder(timestamp), "Cannot open folder."
    
    project_manager.ImportProject(project_path)
    project_manager.LoadProject(project_name)

    resolve_project = project_manager.GetCurrentProject()
    assert resolve_project, "Failed to get Project"

    return resolve_project


def _set_timeline(project, timeline_name):
    for i in range(int(project.GetTimelineCount())):  # GetTimelineCount returns float...
        timeline = project.GetTimelineByIndex(i + 1)  # index starts from 1
        # print(timeline.GetName())
        if timeline.GetName() == timeline_name:
            print( "Setting current timeline to", timeline.GetName())
            assert project.SetCurrentTimeline(timeline), "Cannot set timeline."


def _setup_render_job(project, formatted_output_path, render_preset=""):
    assert project.DeleteAllRenderJobs(), "Cannot delete render jobs..."

    print("Loading render preset [{}]".format(render_preset))
    assert project.LoadRenderPreset(render_preset), "Cannot set render_preset [{}]".format(render_preset)

    render_settings = {
        # "SelectAllFrames": ?,
        # "MarkIn": ?,
        # "MarkOut": ?,
        "TargetDir": os.path.dirname(formatted_output_path),
        "CustomName": os.path.basename(formatted_output_path)
    }
    assert project.SetRenderSettings(render_settings), "Cannot set render settings..."

    jobId = project.AddRenderJob()

    assert jobId, "Cannot add render job..."

    return jobId


def _start_render(project, jobId):
    assert project.StartRendering(jobId)
    while project.IsRenderingInProgress():
        # print(project.GetRenderJobStatus(jobId))

        status = project.GetRenderJobStatus(jobId)

        if status.get("CompletionPercentage"):
            print("Progress: {}%".format(status.get("CompletionPercentage")))

        # sys.stdout.flush()
        time.sleep(1)
    print(project.GetRenderJobStatus(jobId))

    if status.get("CompletionPercentage"):
        print("Progress: {}%".format(status.get("CompletionPercentage")))


if __name__ == '__main__':
    main()
