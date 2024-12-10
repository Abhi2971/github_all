import os
import shutil
from openpyxl import Workbook
from datetime import datetime
from git import Repo

# Base directory for managing repositories
TARGET_FOLDER = "./TeamTaskRepo"

# Excel file paths
EXCEL_FILE = "StudentRepositoryData.xlsx"
COMMIT_EXCEL_FILE = "CommitDetails.xlsx"

# Repository URLs and configurations
ABHI_REPO_URL = "https://github.com/Abhi297127/java.git"
MEMBER_REPOS = {
    "Abhishek_Shelke": "https://github.com/Abhi2971/APJFSA_N.git",
    "Aditi_Sandbhor": "https://github.com/AditiSandbhor/JavaRepository",
    "Bhushan_Ingle": "https://github.com/Bhushan1828/JavaRepo.git",
    "Atharv_Patekar": "https://github.com/Athruu07/JavaProgram_C9550-.git",
    "Shubhangi_Pawar": "https://github.com/shubhangi-Pawar-bit/APJFSA_N_ANP-C9550",
    "Pooja_Yadav": "https://github.com/pjyadav321/APJFSA_N.git",
    "Aman_Bisen": "https://github.com/amanbisen45/APJFSA_N",
    "Bindiya_Shetty": "https://github.com/BindiyaShetty18/Anudip-F__Java-Programs.git",
    "Rutik_Danavale": "https://github.com/RutikDanavale/Lab.git",
    "Sanket_Chivhe": "https://github.com/Sanketchivhe/Javalab.git",
    "Lubna_Kazi": "https://github.com/Lubnakazi27/AJPFSA",
    "Sakshi_Khopade": "https://github.com/sakshikhopade/APJFSA_N.git",
    "Prachi_Salunkhe": "https://github.com/prachi-salunkhe2003/APJFSA_N.git",
    "Sanket_Bhoite": "https://github.com/SB2157/APJFSA_N.git",
    "Aditi_Zanje": "https://github.com/Aditizanje/APJFSA_N.git",
    "Vaishnavi_Mane": "https://github.com/v29m01/AJPF",
    "Dipak_Kondhalkar": "https://github.com/dipakkondhalkar/APJFSA_N.git",
    "Omkar_Kudale": "https://github.com/omkar17k/omkar17k.git",
    "Mangesh_Kolapkar": "https://github.com/mangeshkolapkar1144/APJFSA.git",
    "Nikita_Karape": "https://github.com/nikitakarape/APJFSA_N.git",
    "Prabhat_Sharma": "https://github.com/PrabhatASharma/APJFSA_N.git",
    "Dhananjay_Ghate": "https://github.com/DhananjayGhate111/ANP-CORE_JAVA.git",
    "Bhuvan_Bhoge": "https://github.com/bhuvanbhoge/APJFSA_N",
    "Shubham_Chavan": "https://github.com/Jarvis1401/Anudip_java_Classes.git",
    "Aniruddha_Vinchurkar": "https://github.com/aniruddhaa7/APJFSA-JavaProgram",
    "Sanskruti_Kale": "https://github.com/sanskruti13/APJFSA_JAVA",
    "Suraj_Yadav": "https://github.com/Surajone8/Anudeep_Java",
    "Pranav_Tamboi": "https://github.com/PranavTamboli/APJFSA_N",
    "Shruti_Chaudhari": "https://github.com/shruti-chaudharii/APJFSA_JAVA_COURSE",
    "Shreyash_Pawar": "https://github.com/shreyashpawar8910/APJFSA_N",
    "Aachal_Akre": "https://github.com/aachalakre001/APJFSA_N",
    "Pravin_Subrav_Gunjal": "https://github.com/Pravin224/ANP_java_classes.git",
    "Kaivalya_Kulkarni": "https://github.com/kaivalyakulkarni123/APJSFA_M.git",
    "Bhushan_Sonje": "https://github.com/Bhushan2211/Java_programming.git",
    "Shriya_Wankhede": "https://github.com/shriyawankhede/core-java.git",
    "Vaibhav_Bhange": "https://github.com/vaibhav-11-hub/Java-programs.git",
    "Aarya_Shinde": "https://github.com/Aarya-Shinde/Lab-Work-.git",
    "Ankush_Mishra": "https://github.com/ankushmishra1/APJFSA_N_.git",
    "Madhuri_Shinde": "https://github.com/Madhurishinde20/Java.git",
    "Siddhant_Mane": "https://github.com/SiddhantM007/Java-Lab-1",
    "Omkar_Gaikwad": "https://github.com/Omkar3321/Java-Course-",
    "Nikhil_Patil": "https://github.com/Nickpatil45/APJFSA_N",
    "Vishruti_Rane": "https://github.com/vishrutirane",
    "Ankita_More": "https://github.com/ANKITAMORE06/APJFSA_N-JAVA.git",
    "Vaishnavi_Deshmukh": "https://github.com/VaishuDeshmukh-2003/APJFSA_N.git"
}

START_DATE = datetime.strptime("2024-10-20", "%Y-%m-%d")
END_DATE = datetime.now()

def delete_git_folder_and_add_placeholder(directory):
    git_dir = os.path.join(directory, ".git")
    if os.path.exists(git_dir):
        shutil.rmtree(git_dir)
    if not os.listdir(directory):
        placeholder_path = os.path.join(directory, ".keep")
        with open(placeholder_path, "w") as placeholder_file:
            placeholder_file.write("")

def count_java_files(directory):
    java_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".java"):
                java_files.append(os.path.join(root, file))
    return java_files

def get_commits_in_date_range_and_verify(repo_path, start_date, end_date):
    commits = []
    java_file_sum = 0
    all_java_files = []  # Use a list to track duplicates
    affected_files = []

    try:
        repo = Repo(repo_path)
        for commit in repo.iter_commits():
            commit_date = datetime.fromtimestamp(commit.committed_date)
            if start_date <= commit_date <= end_date:
                commit_java_files = []  # Use a list for tracking Java files (with duplicates)

                # If commit has parents, compare with the parent diff
                if commit.parents:
                    diff = commit.diff(commit.parents[0], create_patch=False)
                    for d in diff:
                        file_name = os.path.basename(d.b_path if d.b_path else d.a_path)
                        if file_name.endswith(".java"):
                            commit_java_files.append(file_name)  # Allow duplicates
                else:
                    # For initial commit (without parent), check all files in the commit
                    for path in repo.tree(commit).traverse():
                        if path.path.endswith(".java"):
                            commit_java_files.append(os.path.basename(path.path))  # Allow duplicates

                # Update totals
                java_file_sum += len(commit_java_files)
                all_java_files.extend(commit_java_files)  # Add duplicates to the list
                affected_files.append(", ".join(commit_java_files))  # Join changed files as a string
                
                # Collect commit details
                commits.append({
                    "date": commit_date.strftime("%Y-%m-%d %H:%M:%S"),
                    "message": commit.message.strip(),
                    "author": commit.author.name,
                    "id": commit.hexsha,
                    "java_file_count": len(commit_java_files),
                    "files": ", ".join(commit_java_files),
                })
    except Exception as e:
        print(f"Error: {str(e)}")  # Print the error message (for debugging)

    return commits, java_file_sum, affected_files
def create_excel_sheet(data, file_path, headers, sheet_name):
    if os.path.exists(file_path):
        os.remove(file_path)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(headers)
    for row in data:
        sheet.append(row)
    workbook.save(file_path)

def clean_target_folder(folder_path):
    if os.path.exists(folder_path):
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if item != ".git":
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.unlink(item_path)
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
        print("Deleted all files except .git")

def process_repositories():
    excel_data = []
    commit_excel_data = []
    failed_repos = []

    for member, repo_url in MEMBER_REPOS.items():
        member_dir = os.path.join(TARGET_FOLDER, member)
        if not os.path.exists(member_dir):
            os.makedirs(member_dir)
            try:
                print(f"Cloning {member}...")
                Repo.clone_from(repo_url, member_dir)
                java_files = count_java_files(member_dir)
                java_count = len(java_files)

                commits, total_java_count_from_commits, uploaded_files = get_commits_in_date_range_and_verify(
                    member_dir, START_DATE, END_DATE
                )
                for i, commit in enumerate(commits):
                    commit_excel_data.append([
                        member,
                        repo_url,
                        commit["date"],
                        commit["author"],
                        commit["message"],
                        commit["java_file_count"],
                        uploaded_files[i],
                        commit["id"],
                    ])
                excel_data.append([member, repo_url, java_count, "Success"])
                delete_git_folder_and_add_placeholder(member_dir)
            except Exception as e:
                failed_repos.append((member, repo_url))
                excel_data.append([member, repo_url, "N/A", "Failed"])
    
    return excel_data, commit_excel_data, failed_repos

def main():
    clean_target_folder(TARGET_FOLDER)

    # Process repositories and generate Excel files
    excel_data, commit_excel_data, failed_repos = process_repositories()
    create_excel_sheet(excel_data, EXCEL_FILE, ["Member Name", "Repository URL", "Java File Count", "Status"], "Repository Data")
    create_excel_sheet(commit_excel_data, COMMIT_EXCEL_FILE, 
                       ["Member Name", "Repository URL", "Commit Date", "Author", "Commit Message","Java File Count", "Uploaded Files","Commit ID"], 
                       "Commits Data")
    
    # Print summary
    if os.path.exists(TARGET_FOLDER) and os.path.isdir(os.path.join(TARGET_FOLDER, ".git")):
        try:
            repo = Repo(TARGET_FOLDER)
            repo.git.add(A=True)
            repo.index.commit("updated all codes")
            origin = repo.remote(name="origin")
            origin.push()
            print("Pushed cleanup changes to GitHub.")
        except Exception as e:
            print(f"Error during Git operations after cleanup: {e}")
    else:
        print(f"Error: {TARGET_FOLDER} is not a Git repository.")

    print("\nSummary:")
    print(f"Total repositories processed: {len(MEMBER_REPOS)}")
    if failed_repos:
        print(f"Failed repositories: {len(failed_repos)}")
        for i, (member, repo_url) in enumerate(failed_repos, 1):
            print(f"{i}. {member}: {repo_url}")

if __name__ == "__main__":
    main()

