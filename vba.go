/*
 *  Copyright (c) 2020, Peter Haag
 *  All rights reserved.
 *  
 *  Redistribution and use in source and binary forms, with or without 
 *  modification, are permitted provided that the following conditions are met:
 *  
 *   * Redistributions of source code must retain the above copyright notice, 
 *     this list of conditions and the following disclaimer.
 *   * Redistributions in binary form must reproduce the above copyright notice, 
 *     this list of conditions and the following disclaimer in the documentation 
 *     and/or other materials provided with the distribution.
 *   * Neither the name of the author nor the names of its contributors may be 
 *     used to endorse or promote products derived from this software without 
 *     specific prior written permission.
 *  
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 *  AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE 
 *  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE 
 *  ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE 
 *  LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR 
 *  CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
 *  SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
 *  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
 *  CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
 *  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE 
 *  POSSIBILITY OF SUCH DAMAGE.
 *  
 */

package vba

import (
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"os"
)

// max 4096 data + 2 header
const VBABuffSize = 4098
type VBAWriter struct {
	wr io.Writer
	compressedBuffer   []byte
	uncompressedBuffer []byte
	cSize int
	uSize int
	marker struct {
		cur   int
		found bool
	}
}

func New() *VBAWriter {
	return new(VBAWriter)
}

func (vba *VBAWriter)init() {
	vba.compressedBuffer   = make([]byte, VBABuffSize)
	vba.uncompressedBuffer = make([]byte, VBABuffSize)
}

func NewWriter(wr io.Writer) *VBAWriter {
	vba := new(VBAWriter)
	vba.wr = wr
	vba.init()
	return vba
}

func (vba *VBAWriter)Close() error {
//	fmt.Printf("vba.Close()\n")
//	fmt.Printf("cSize: %d, uSize: %d\n", vba.cSize, vba.uSize)
	return nil
}

func (vba *VBAWriter)uncompress() (int, error) {

	total := 0
	for vba.cSize > 0 {
		if processed, err := vba.uncompressChunk(); err != nil {
			return total + processed, err
		} else {
			if processed == 0 {
				// not enough data in buffer - fiil with next write
				break
			}
			total += processed
			if vba.uSize != 0 {
				// write uncompressed block to underlaying writer
				n, err := vba.wr.Write(vba.uncompressedBuffer[:vba.uSize])
				if err != nil {
					return total, err
				}
				if n != vba.uSize {
					return total, fmt.Errorf("Short write to output stream\n")
				}
				// reset buffer
				vba.uSize = 0
			}
			copySize := vba.cSize - processed
			if copySize < 0 {
				return total, fmt.Errorf("Uncompress error: Negativ copy size: %d\n", copySize)
			}
			if copySize > 0 {
				// shift remaining bytes
				copy(vba.compressedBuffer[0:copySize], vba.compressedBuffer[processed:vba.cSize])
			}
			vba.cSize = copySize
		}
	}
	return total, nil

} // End of uncompress

func (vba *VBAWriter)Write(b []byte) (int, error) {
	var offset int
	var hasVBA bool
	if !vba.marker.found {
		if offset, hasVBA = vba.carvVBA(b); !hasVBA {
			return len(b), nil
		} // else fall through
	}

	// VBA marker true - loop over buffer
	buffAvail := len(b) - offset
	vbaAvail  := VBABuffSize - vba.cSize
	total := 0
	for buffAvail > 0 {
		if vbaAvail < buffAvail {
			copy(vba.compressedBuffer[vba.cSize:], b[offset:offset+vbaAvail])
			offset += vbaAvail
			vba.cSize += vbaAvail
		} else {
			copy(vba.compressedBuffer[vba.cSize:vba.cSize+buffAvail], b[offset:])
			offset += buffAvail
			vba.cSize += buffAvail
		}

		if processed, err := vba.uncompress(); err != nil {
			return total + processed, err
		}

		vbaAvail = VBABuffSize - vba.cSize
		buffAvail = len(b) - offset
	}

	return len(b), nil
}

// carv for VBA signature in stream
// returns true if signature found 
func (vba *VBAWriter)carvVBA(buff []byte) (int, bool) {

    marker	:= []byte("\x01..\x00Attribut")
	mlen	:= len(marker)

	for i, b := range buff {
		if vba.marker.cur == 1 || vba.marker.cur == 2 {
			// unconditionally add these 2 bytes
			vba.compressedBuffer[vba.cSize] = b
			vba.marker.cur++
			vba.cSize++
		} else {
			if b == marker[vba.marker.cur] {
				vba.compressedBuffer[vba.cSize] = b
				vba.marker.cur++
				if vba.marker.cur != 1 {
					// skip signature byte \x01 in uncompressed buffer
					vba.cSize++
				}
			} else {
				vba.marker.cur = 0
				vba.cSize = 0
			}
		}

		if vba.marker.cur == mlen {
			vba.marker.found = true
			return i+1, true
		}
	}

	return 0, false

} // End of CarvVBA

func copyTokenHelp(current uint, chunkStart uint) (uint16, uint16, uint16, uint16) {

	difference := current - chunkStart
	b := math.Ceil(math.Log2(float64(difference)))
	bitCount := uint16(math.Max(b, 4.0))
	lengthMask := uint16(0xFFFF >> bitCount)
	offsetMask := ^lengthMask
	maximumLength := uint16((0xFFFF >> bitCount) + 3)
	return lengthMask, offsetMask, bitCount, maximumLength

} // End of copyTokenHelp

func (vba *VBAWriter)uncompressChunk() (int, error) {

	CompressedHeader := uint16(binary.LittleEndian.Uint16(
		vba.compressedBuffer[0:2]))
	vba.uSize = 0

	CompressedChunkFlag := (CompressedHeader >> 15) & 0x1
	CompressedChunkSignature := (CompressedHeader >> 12) & 0x7
	CompressedChunkSize := (CompressedHeader & 0x0FFF) + 3

	if CompressedChunkSignature != 0x3 { // b011 => 0x3
		return 0, fmt.Errorf("Chunk signature missmatch")
	}

	CompressedEnd := int(CompressedChunkSize)
	CompressedCurrent := 2
	if CompressedEnd > vba.cSize {
		// not enough data in buffer
		return 0, nil
	}

	if CompressedChunkFlag == 0 {
		if CompressedChunkSize != 4098 {
			return 0, fmt.Errorf("Uncompressed chunk size error")
		}
		copy(vba.uncompressedBuffer[vba.uSize:vba.uSize+4096],
			 vba.compressedBuffer[CompressedCurrent : CompressedCurrent+4096])
		CompressedCurrent += 4096
		vba.uSize += 4096
	} else {
		decompressedChunkStart := vba.uSize
		for CompressedCurrent < CompressedEnd {
			flagByte := vba.compressedBuffer[CompressedCurrent]
			CompressedCurrent++
			var bitIndex uint
			for bitIndex = 0; bitIndex < 8; bitIndex++ {
				if CompressedCurrent < CompressedEnd {
					flagBit := (flagByte >> bitIndex) & 0x1
					if flagBit == 0 {
						vba.uncompressedBuffer[vba.uSize] = vba.compressedBuffer[CompressedCurrent]
						vba.uSize++
						CompressedCurrent++
					} else {
						var copyToken uint16
						copyToken = binary.LittleEndian.Uint16(vba.compressedBuffer[CompressedCurrent : CompressedCurrent+2])
						lengthMask, offsetMask, bitCount, _ := copyTokenHelp(uint(vba.uSize), uint(decompressedChunkStart))
						length := int((copyToken & lengthMask) + 3)
						temp1 := copyToken & offsetMask
						temp2 := 16 - bitCount
						offset := int((temp1 >> temp2) + 1)
						copySource := vba.uSize - offset
						for index := copySource; index < copySource+length; index++ {
							vba.uncompressedBuffer[vba.uSize] = vba.uncompressedBuffer[index]
							vba.uSize++
						}
						CompressedCurrent += 2
					}
				}
			}
		}
	}

	return CompressedCurrent, nil
} // End of uncompressChunk

func DecompressFile(inFile, outFile string) (bool, error) {

	var rfd, wfd *os.File
	var err error
	if inFile == "-" {
		rfd = os.Stdin
	} else {
		if finfo, err := os.Stat(inFile); err != nil {
			return false, err
		} else if !finfo.Mode().IsRegular() {
			return false, fmt.Errorf("%s is not a regualar file", inFile)
		}

		rfd, err = os.Open(inFile)
		if err != nil {
			return false, err
		}
	}

	if outFile == "-" {
		wfd = os.Stdout
	} else {
		wfd, err = os.OpenFile(outFile, os.O_RDWR|os.O_CREATE, 0755)
		if err != nil {
			return false, err
		}
	}

	vba := NewWriter(wfd)
	_, err = io.Copy(vba, rfd)

	vba.Close()
	if inFile != "-" {
		rfd.Close()
	}
	if outFile != "-" {
		wfd.Close()
	}
	if err != nil {
		return false, fmt.Errorf("VBA: %v\n", err)
	}

	return true, nil

} // End of DecompressFile
