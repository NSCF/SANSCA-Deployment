Step-by-step installation guide

# Server Setup

- 16384MB RAM
- 127GB SSD
- 4 vCPU
- Ubuntu 22.04 LTS running on Windows Server 2016 Hyper-V

# Installation Guide

https://github.com/AtlasOfLivingAustralia/ala-install/tree/master

## Notes

- Unable to perform `apt-get install python-dev`. The command auto-completes to: `python-dev-is-python3`
- Unable to run `sudo pip install setuptools`
- Installed `sudo apt -get install python3-pip`
- Successfully completed `pip install ansible==9.5.1 ansible-core==2.16.6`
- Installed Docker to use **LA-toolkit**

```bash
- # Add Docker's official GPG key:
sudo apt-get update
sudo apt-get install ca-certificates curl
sudo install -m 0755 -d /etc/apt/keyrings
sudo curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
sudo chmod a+r /etc/apt/keyrings/docker.asc

# Add the repository to Apt sources:
echo \
  "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu \
  $(. /etc/os-release && echo "${UBUNTU_CODENAME:-$VERSION_CODENAME}") stable" | \
  sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt-get update
```

```bash
sudo apt-get install docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin -y
```

```bash
sudo usermod -aG docker $USER
```

```bash
sudo apt-get install docker-compose -y
sudo reboot
```

```bash
git clone https://github.com/living-atlases/la-toolkit.git
cd la-toolkit
```

## Important

For the LA-Toolkit to successfully connect and validate to the server, to be able to perform the necessary Ansible playbooks. You must ensure that the username set in the LA-Toolkit UI is defined in `visudo` with `NOPASSWD` being defined.
- In `visudo` you must add the user account as such: `{USERNAME} ALL=(ALL) NOPASSWD:ALL`
