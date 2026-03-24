# Certificate Management

## ManageCerts.ps1

The `ManageCerts.ps1` script is your entry point for managing certs. This is intended for your own local network or lab setup, and is _not_ intended for use within an organization.

### Parameters

1. `TargetDirectory` - Allows you to specify the directory you want to run the script execution in. If you don't specify it the default is the script root, _not_ the current directory. So if you place these in your path, you still need to specify `./`.
2. `Devices` - Allows you to specify one, or several, devices that you want to perform updates on. Useful when something like a firmware update overwrites your certificates and you need to restore them for only specific devices.

### Operation

#### Initial Configuration

When you first run the script in a directory there will be no configuration, so you'll be guided through initial configuration steps. Here is the general step-by-step process you'll be prompted through.

1. You'll be prompted for a country code for your Cert, generally you should use the code for the country your devices reside in. This is for the Root Certificate Authority certificate. Examples would be `US`, `DE`, `CA`, etc.
2. Next you'll be prompted for the organizational unit name. Although this can be anything you want, typically keeping it descriptive and in alignment with the Fully Qualified Domain Names you'll be using later is helpful. So, if you plan on having FQDNs like `proc1.local` you could name your org unit `Local`, or you could do your last name, or your initials, or anything that you would use to identify your local network via DNS.
3. The CA certificate file name is next and is typically something like `local-ca`. This will be located in a subdirectory titled `local-ca`.
4. Next you'll be prompted for the default password to use for Crestron device certificates. Some devices require certificates have a key, so skipping this step isn't optional if you're loading to Crestron devices.
5. Now you'll be prompted to add devices. If you press enter when prompted for the FQDN for a device it will finish the setup and move to processing the certs.
6. Enter your first FQDN or simple hostname in the form `host.network.com`, `host.local`, or just `hostname`. This is used as the certificate Common Name and Subject Alternative Name, and as the name for the device's certificate directory.
7. Next enter the IP Address of the device. If you have multiple IPs it is accessible from (like for a router w/ multiple networks, or a Crestron processor with a control subnet, or a device with multiple NICs that have different IPs for different purposes) you can enter them all using commas to separate them. `IP1,IP2,IP3`
8. Optionally, enter a **connection address** to use when connecting to the device for certificate deployment. This is useful when the FQDN does not resolve on your network and you need to connect via an IP address instead. If left blank the FQDN will be used for connection (Crestron devices) or the first IP address will be used (all other device types). You can also set this to any hostname or IP that is reachable on your network.
9. Next you'll be prompted for a username. At the time of this documentation all methods for deploying certs require usernames, except for the TrueNAS, as it uses an API key instead.
10. Next up, the password. Because I wrote this script to live on local machines and not be saved to the cloud or anything, it's pretty darn insecure to keep your passwords here, but it's what I required, because I don't want to have to enter the password manually for every device I deploy a cert to. I might try to fix this in the future, maybe not. If you need a password for the deployment method, enter it here. The Unifi method potentially allows authenticating with a local SSH key, which requires providing the path to. You have to manually edit your config to enter this information; see the example config for details.
11. Next choose what kind of update type. Currently there are the following types:
    1. `Crestron3` - For creating and deploying certs compatible with Crestron Series 3 processors.
    2. `Crestron4` - For creating and deploying certs compatible with Crestron Series 4 processors.
    3. `CrestronTP60Series` - For creating and deploying certs compatible with Crestron 60 Series touchpanels.
    4. `CrestronTP70Series` - For creating and deploying certs comptaible with Crestron 70 Series touchpanels. This may also work with 80 series panels, but it hasn't been tested.
    5. `TrueNAS` - For deploying certs to a TrueNAS Scale server and configuring the UI to use that cert. Requires configuring an API key on the server and adding it to your config, which must be done manually at the moment.
    6. `UniFi` - For deploying certs to a UNIFI device like a Dream Machine Pro, or similar. Requires enabling SSH access on the device.
    7. `SCP` - For deploying certs to a directory on a specified device. Allows deploying any of the following types of cert files: `RootCA`, `PFX`, `PEM`, `CRT`, `KEY`, which must be manually entered into a hash table in the configuration. See the example config for examples.
12. You'll be prompted again for a new device; if you're done, just hit enter, otherwise repeat the steps again. Once you hit enter here the script will start generating and updating certs.
13. Lastly you will be prompted for the password for the ROOT CA password. This is never stored in the config and is required to be entered everytime you run the script. DO NOT LOSE THIS PASSWORD, as it is required for properly generating your certs.

If you need to manually edit the config at all, you will have to do so now, using the example config for samples.

### Configuration Reference

Each device entry in `CertificatesToManage` supports the following fields:

- `FQDN` *(required)* - The hostname or fully qualified domain name used as the certificate's Common Name and Subject Alternative Name. Can be a simple hostname (e.g., `mydevice`) or an FQDN (e.g., `mydevice.example.com`).
- `IPAddress` *(required)* - An array of IP addresses included in the certificate's Subject Alternative Names. The first IP is used as the default connection address for non-Crestron device types.
- `ConnectionAddress` *(optional)* - The address used for network connections when deploying certificates. If omitted, Crestron devices connect via the FQDN and all other device types connect via the first IP address. Set this to an IP address or alternate hostname when the FQDN does not resolve on your network.
- `Username` - Username for device authentication (not required for TrueNAS).
- `Password` - Password for device authentication (not required for TrueNAS or when using `SshKeyPath`).
- `UpdateType` *(required)* - The deployment method to use (see update types listed above).
- `ApiKey` - API key for TrueNAS authentication.
- `SshKeyPath` - Path to an SSH private key file for key-based authentication (UniFi and SCP).
- `FileMappings` - A map of certificate type keys to remote file paths (SCP only).

#### Subsequent Execution

Run the script again. If a config file exists you'll skip the configuration and go straight to generating and deploying certs.

## Other Scripts

The other scripts `Create-ClientCert.ps1` and `Create-RootCert.ps1` are invoked from within the `ManageCerts.ps1` script and aren't really intended to be run stand-alone.