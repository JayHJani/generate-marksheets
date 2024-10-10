import * as React from 'react';
import Button from '@mui/material/Button';
import * as XLSX from 'xlsx';
import { Typography } from '@mui/material';
import AttachFileIcon from '@mui/icons-material/AttachFile';
import FileUploadIcon from '@mui/icons-material/FileUpload';

interface StudentData {
    rollNumber: string
    id: number,
    name: string,
    result: number,
}

const AcceptedFileType = {
    Excel: '.xlsx',
};

export function UploadFileBtn() {
    const fileRef = React.useRef<HTMLInputElement>(null);
    const acceptedFormats = AcceptedFileType.Excel;

    const [selectedFiles, setSelectedFiles] = React.useState<any>();

    // Utility function to format the current date
    const getCurrentDateTimeString = (): string => {
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0'); // Months are zero-based
        const dd = String(now.getDate()).padStart(2, '0');
        const hh = String(now.getHours()).padStart(2, '0');
        const min = String(now.getMinutes()).padStart(2, '0');
        const sec = String(now.getSeconds()).padStart(2, '0');

        return `${yyyy}-${mm}-${dd}_${hh}-${min}-${sec}`; // Format: YYYY-MM-DD_HH-MM-SS
    };

    const handleFileSelect = (event: any) => {
        setSelectedFiles(event?.target?.files?.[0]);
    };

    const onUpload = async () => {
        const originalFileName = selectedFiles?.name
        const data = await selectedFiles?.arrayBuffer()
        const workbook = XLSX.read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: "",
        });

        // parse data
        const otherDetails = jsonData.slice(0, 4)
        const studentDetails = jsonData.slice(4)

        let map: Map<string, StudentData[]> = new Map()

        studentDetails.map((data: unknown) => {
            if (Array.isArray(data) && data.length >= 4) {
                let studentData: StudentData = {
                    rollNumber: data[0],
                    id: data[1],
                    name: data[2],
                    result: data[3],
                };

                const division = data[0][data[0].length - 1]
                const currStudents = map.get(division)
                if (currStudents !== undefined) {
                    map.set(division, [...currStudents, studentData])
                } else {
                    map.set(division, [studentData])
                }
            }
        })

        const allDivisions = map.keys()
        const allDivisionsArray = Array.from(allDivisions)
        allDivisionsArray.sort()

        const newWorkBook = XLSX.utils.book_new()
        allDivisionsArray.map(division => {
            const studentDataForGivenDivision = map.get(division)
            if (studentDataForGivenDivision !== undefined) {
                const worksheet = XLSX.utils.json_to_sheet(studentDataForGivenDivision)
                XLSX.utils.book_append_sheet(newWorkBook, worksheet, division);
            }
        })
        const currentDate = getCurrentDateTimeString();
        XLSX.writeFile(newWorkBook, `Internal_Marksheet_${currentDate}.xlsx`);
    };

    return (
        <>
            <div style={{ display: 'flex', flexDirection: 'row', justifyContent: 'center', gap: "20px", marginTop: '40vh'}}>
                <input
                    ref={fileRef}
                    hidden
                    type="file"
                    accept={acceptedFormats}
                    onChange={handleFileSelect}
                />
                <Button
                    variant="outlined"
                    component="label"
                    style={{ textTransform: 'none' }}
                    onClick={() => fileRef.current?.click()}
                >
                    <AttachFileIcon fontSize='small' />
                    Choose file to upload
                </Button>
                <Button
                    color="primary"
                    variant='contained'
                    disabled={!selectedFiles}
                    style={{ textTransform: 'none' }}
                    onClick={onUpload}
                >
                    <FileUploadIcon />
                    Upload
                </Button>
            </div>

            {
                selectedFiles?.name ? (
                    <div style={{ display: 'flex', flexDirection: 'row', justifyContent: 'center', gap: "20px", paddingTop: "10px" }}>
                        <Typography>Selected File : </Typography>
                        <Typography variant='subtitle2'>{selectedFiles?.name} </Typography>
                    </div>
                ) : <></>
            }


        </>
    );
}