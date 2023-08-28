import os
import subprocess

def install_conda_environment(env_name):
    subprocess.run(['conda', 'create', '-n', env_name, 'python=3.8.17', '-y'])

def install_dependencies():
    subprocess.run(['conda', 'install', '-n', 'projects', 'requests', 'beautifulsoup4', 'selenium', 'pillow', 'opencv', 'openpyxl', 'docx2pdf', 'comtypes', '-y'])

def main():
    env_name = 'projects'

    print("Installing conda environment...")
    install_conda_environment(env_name)

    print("Activating conda environment...")
    activate_cmd = f"conda activate {env_name}"
    subprocess.run(activate_cmd, shell=True, executable="/bin/bash")

    print("Installing project dependencies...")
    install_dependencies()

    print("Setup complete!")

if __name__ == '__main__':
    main()
