import os
import subprocess
import sys

def run_cmd(command, cwd=None):
    print(f"\nRunning command: {command} in {cwd or 'root'}")
    # On Windows powershell/cmd running script is disabled sometimes, we use shell=True and call cmd directly
    result = subprocess.run(command, shell=True, cwd=cwd)
    if result.returncode != 0:
        print(f"Command failed with exit code: {result.returncode}")
        return False
    return True

def main():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. Install root dependencies
    print("=== Step 1: Installing root dependencies ===")
    if not run_cmd("npm.cmd install", root_dir):
        sys.exit(1)
        
    # 2. Install backend dependencies
    print("=== Step 2: Installing NestJS backend dependencies ===")
    backend_dir = os.path.join(root_dir, "backend")
    if not run_cmd("npm.cmd install", backend_dir):
        sys.exit(1)
        
    # 3. Install frontend dependencies
    print("=== Step 3: Installing Next.js frontend dependencies ===")
    frontend_dir = os.path.join(root_dir, "frontend")
    if not run_cmd("npm.cmd install", frontend_dir):
        sys.exit(1)
        
    # 4. Generate Prisma Client and SQLite Database
    print("=== Step 4: Initializing Prisma Client & Database ===")
    if not run_cmd("npx.cmd prisma generate", backend_dir):
        sys.exit(1)
        
    if not run_cmd("npx.cmd prisma db push", backend_dir):
        sys.exit(1)
        
    print("\n=== Setup completed successfully! ===")
    print("To start the development servers concurrently, run:")
    print("  npm.cmd run dev")
    print("This will launch: ")
    # Frontend port is 3000, Backend port is 8001
    print("  - Frontend: http://localhost:3000")
    print("  - Backend: http://localhost:8001")

if __name__ == "__main__":
    main()
