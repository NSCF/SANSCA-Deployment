## How the SHA256_checksum_tool works
The script relies on Windows’ built-in **certutil** command to calculate SHA-256 hashes. If certutil is not available, the script cannot generate or verify checksums, so the check ensures the required tool is present before proceeding. **certutil** should already be included in your Windows installation.

## Drag and Drop
* Drag a folder onto the script and wait for popup.
* Type G for generate, V for verify.
* A SHA256SUMS.txt will be generated in the same folder as the files being hashed.

## Command-line usage
```
SHA256_checksum_tool.bat generate "D:\MasterFolder"
```
* A SHA256SUMS.txt will be generated in the same folder as the files being hashed.
```
SHA256_checksum_tool.bat verify "D:\MasterFolder"
````
* Reads the SHA256SUMS.txt file and checks if it is the same.

## Folder-level SHA256 hash
A single SHA-256 checksum computed from the manifest file, representing the entire folder’s contents. It allows quick verification that all files remain unaltered without recalculating each file’s hash individually.
* Can be found at the bottom of SHA256SUMS.txt

