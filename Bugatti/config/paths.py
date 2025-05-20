import sys
import os

def get_team_member():
    # Check command line arguments first
    if len(sys.argv) > 1 and not sys.argv[1].endswith(".py"):
        return sys.argv[1]
    
    # Check environment variable
    if "TEAM_MEMBER" in os.environ:
        return os.environ["TEAM_MEMBER"]
    
    # Default to Tima if not specified
    return "Tima"

# Get team member
TEAM_MEMBER = get_team_member()

# Create standard paths used across scripts
def get_paths():
    base_dir = "J:\\SPMS_Registration_Structured"
    member_dir = os.path.join(base_dir, "Team Members", TEAM_MEMBER)
    
    paths = {
        "base_dir": base_dir,
        "member_dir": member_dir,
        "pet_forms": os.path.join(member_dir, "PetForms"),
        "uploads": os.path.join(member_dir, "Uploads"),
        "scripts_dir": os.path.join(base_dir, "Bugatti")
    }
    
    # Create directories if they don't exist
    for path in paths.values():
        if path != base_dir and not os.path.isdir(path) and not path.endswith(".py"):
            os.makedirs(path, exist_ok=True)
    
    return paths

# Paths dictionary to use in scripts
PATHS = get_paths()