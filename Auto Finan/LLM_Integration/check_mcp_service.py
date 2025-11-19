"""
Quick check script for MCP service status
Usage: python check_mcp_service.py [--url http://192.168.137.133:3030]
"""

import requests
import sys
import socket

def check_port(host, port, timeout=3):
    """Check if port is open"""
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        result = sock.connect_ex((host, port))
        sock.close()
        return result == 0
    except Exception:
        return False

def main():
    url = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:3030"
    
    # Parse host and port
    if "://" in url:
        host = url.split("://")[1].split(":")[0].split("/")[0]
    else:
        host = "localhost"
    
    port = 3030
    if ":" in url.split("://")[-1]:
        try:
            port = int(url.split(":")[-1].split("/")[0])
        except:
            pass
    
    print(f"Checking MCP service at {url}")
    print(f"Host: {host}, Port: {port}")
    print()
    
    # Check port
    print("1. Checking port...")
    if check_port(host, port):
        print(f"   [OK] Port {port} is open")
    else:
        print(f"   [FAIL] Port {port} is NOT open")
        print()
        print("Possible causes:")
        print("  - Service is not running")
        print("  - Firewall is blocking the connection")
        print("  - Service is listening on localhost instead of 0.0.0.0")
        if host != "localhost" and host != "127.0.0.1":
            print(f"  - Remote host {host} is unreachable")
        print()
        print("Solutions:")
        print("  1. Start the service: start_mcp_gateway.bat")
        print("  2. Check if service listens on 0.0.0.0:3030")
        print("  3. Check Windows Firewall settings")
        return False
    print()
    
    # Check health endpoint
    print("2. Checking health endpoint...")
    try:
        response = requests.get(f"{url}/health", timeout=5)
        if response.status_code == 200:
            print(f"   [OK] Health check passed")
            print(f"   Response: {response.json()}")
        else:
            print(f"   [FAIL] Health check failed: {response.status_code}")
    except Exception as e:
        print(f"   [FAIL] Health check failed: {e}")
    print()
    
    # Check execute endpoint
    print("3. Checking execute endpoint...")
    try:
        test_data = {"prompt": "test"}
        response = requests.post(
            f"{url}/mcp/execute",
            json=test_data,
            timeout=10
        )
        if response.status_code == 200:
            print(f"   [OK] Execute endpoint is working")
        else:
            print(f"   [FAIL] Execute endpoint returned: {response.status_code}")
    except Exception as e:
        print(f"   [FAIL] Execute endpoint failed: {e}")
    print()
    
    print("Check completed!")
    return True

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\nInterrupted by user")
        sys.exit(1)

