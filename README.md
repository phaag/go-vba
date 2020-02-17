# go-vba

This package implements the VBA decompression code in Go. It
follows the Golang standards for Reader/Writer model.

## Example:

```go
import {
  "github/phaag/ole-vba"
)
  
rfd, _ = os.OpenFile(inFile)
wfd, _ = os.OpenFile(outFile, os.O_RDWR|os.O_CREATE, 0755)
vba := vba.NewWriter(wfd)
_, err = io.Copy(vba, rfd)
vba.Close()
```


If you prefer a function call, you may use the integrated wrapper:

```go
func DecompressFile(inFile, outFile string) (bool, error)
```



The code follows basically olevba from oletools. The outFile is a simple .txt file containing the decompressed VBA code.

