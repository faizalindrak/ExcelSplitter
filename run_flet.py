#!/usr/bin/env python3
"""
Simple script to run the modern Flet-based Excel Splitter
"""

import sys
import subprocess
from pathlib import Path

def install_requirements():
    """Install requirements if needed"""
    try:
        import flet
        print("✅ Flet is already installed")
        return True
    except ImportError:
        print("📦 Installing Flet and dependencies...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "flet>=0.24.0"])
            print("✅ Flet installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("❌ Failed to install Flet")
            return False

def main():
    """Main entry point"""
    print("🚀 Starting Modern Excel Splitter (Flet)")
    print("=" * 50)
    
    # Check if main_flet.py exists
    flet_app = Path("main_flet.py")
    if not flet_app.exists():
        print(f"❌ {flet_app} not found in current directory")
        return 1
    
    # Install requirements if needed
    if not install_requirements():
        return 1
    
    # Run the application
    try:
        print("🎨 Launching modern UI...")
        subprocess.run([sys.executable, str(flet_app)])
        print("👋 Application closed")
        return 0
    except KeyboardInterrupt:
        print("\n👋 Application interrupted by user")
        return 0
    except Exception as e:
        print(f"❌ Error running application: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())