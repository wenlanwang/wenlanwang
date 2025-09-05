
import React, { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';

declare const saveAs: any;
declare const PizZip: any;

interface Param {
    report_name: string;
    parm_name: string;
    SQL: string;
    report_date: string;
}

const App: React.FC = () => {
    const [templateFile, setTemplateFile] = useState<File | null>(null);
    const [dbFile, setDbFile] = useState<File | null>(null);
    const [reportDate, setReportDate] = useState<string>('');
    const [error, setError] = useState<string>('');
    const [isSettingsOpen, setIsSettingsOpen] = useState<boolean>(false);
    const [isHelpOpen, setIsHelpOpen] = useState<boolean>(false);
    const [params, setParams] = useState<Param[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [filesUploaded, setFilesUploaded] = useState<boolean>(false);

    useEffect(() => {
        const now = new Date();
        const year = now.getFullYear();
        const month = (now.getMonth() + 1).toString().padStart(2, '0');
        setReportDate(`${year}-${month}`);
    }, []);

    const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>, fileType: 'db' | 'template') => {
        const file = e.target.files?.[0];
        if (!file) return;

        if (fileType === 'db') {
            setDbFile(file);
        } else if (fileType === 'template') {
            setTemplateFile(file);
        }
    };
    
    useEffect(() => {
        const uploadFiles = async () => {
            if (templateFile && dbFile) {
                const formData = new FormData();
                formData.append('template', templateFile);
                formData.append('db', dbFile);
                
                setFilesUploaded(false); // Reset on new upload
                setError('');

                try {
                    const response = await fetch('/api/upload', {
                        method: 'POST',
                        body: formData,
                    });
                    const result = await response.json();
                    if (!response.ok) {
                        throw new Error(result.error || 'File upload failed');
                    }
                    setFilesUploaded(true);
                } catch (err: any) {
                    setError(`File upload failed: ${err.message}`);
                    setFilesUploaded(false);
                }
            }
        };
        uploadFiles();
    }, [templateFile, dbFile]);


    const fetchParams = async () => {
        if (!filesUploaded) {
            setError("Please upload files before accessing settings.");
            return;
        }
        try {
            const response = await fetch('/api/params');
            const data = await response.json();
            if (!response.ok) {
                throw new Error(data.error || 'Failed to fetch params');
            }
            setParams(data);
        } catch (err: any) {
            setError(err.message);
        }
    };

    const handleSettingsOpen = () => {
        fetchParams();
        setIsSettingsOpen(true);
    };

    const handleSettingsClose = () => {
        setIsSettingsOpen(false);
    };

    const handleHelpOpen = () => {
        setIsHelpOpen(true);
    };

    const handleHelpClose = () => {
        setIsHelpOpen(false);
    };

    const handleParamChange = (index: number, field: keyof Param, value: string) => {
        const newParams = [...params];
        newParams[index] = { ...newParams[index], [field]: value };
        setParams(newParams);
    };

    const handleAddParam = () => {
        setParams([...params, { report_name: 'new_report', parm_name: 'new_param', SQL: 'SELECT ...', report_date: reportDate }]);
    };
    
    const handleDeleteParam = (index: number) => {
        setParams(params.filter((_, i) => i !== index));
    };

    const handleSaveChanges = async () => {
        try {
            const response = await fetch('/api/params', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(params),
            });
            const result = await response.json();
            if (!response.ok) {
                throw new Error(result.error || 'Failed to save changes');
            }
            handleSettingsClose();
        } catch (err: any) {
            setError(`Failed to save changes: ${err.message}`);
        }
    };

    const handleGenerate = async () => {
        if (!filesUploaded) {
            setError('Please ensure both template and database files are uploaded successfully.');
            return;
        }

        setError('');
        setIsLoading(true);

        try {
            const response = await fetch('/api/generate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ reportDate }),
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.error || `Report generation failed with status: ${response.status}`);
            }

            const blob = await response.blob();
            const contentDisposition = response.headers.get('Content-Disposition');
            let fileName = 'report.docx';
            if (contentDisposition) {
                const fileNameMatch = contentDisposition.match(/filename="(.+)"/);
                if (fileNameMatch && fileNameMatch.length === 2) {
                    fileName = fileNameMatch[1];
                }
            }
            saveAs(blob, fileName);

        } catch (err: any) {
            setError(`Report generation failed: ${err.message}`);
        } finally {
            setIsLoading(false);
        }
    };
    
    const handleDownloadSampleTemplate = () => {
        if (typeof PizZip === 'undefined' || typeof saveAs === 'undefined') {
            setError('Required libraries (PizZip, FileSaver) are not loaded.');
            return;
        }

        try {
            const zip = new PizZip();

            const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:t>Monthly Performance Review: [$parm_month]</w:t></w:r></w:p>
                <w:p><w:r><w:t></w:t></w:r></w:p>
                <w:p><w:r><w:t>Total Sales: [$total_sales]</w:t></w:r></w:p>
                <w:p><w:r><w:t>Top Product: [$top_product]</w:t></w:r></w:p>
                <w:p><w:r><w:t></w:t></w:r></w:p>
                <w:p><w:r><w:t>This sample demonstrates how to use parameters. Replace 'total_sales' and 'top_product' with parameters defined in your settings.</w:t></w:r></w:p>
              </w:body>
            </w:document>`;

            const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>`;

            const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="xml" ContentType="application/xml" />
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
            </Types>`;

            zip.file("[Content_Types].xml", contentTypesXml);
            zip.file("_rels/.rels", relsXml);
            zip.file("word/document.xml", docXml);

            const blob = zip.generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });

            saveAs(blob, "Sample_Report_Template.docx");

        } catch (err: any) {
            console.error(err);
            setError(`Failed to generate sample template. Details: ${err.message}`);
        }
    };


    return (
        <div className="container">
            <h1>Automated Report Generator</h1>

            {error && (
                <div className="error-notification">
                    <span>{error}</span>
                    <button onClick={() => setError('')} className="close-button">&times;</button>
                </div>
            )}

            <div className="input-group">
                <label htmlFor="template-upload">1. Upload Word Report Template (.docx)</label>
                <input id="template-upload" type="file" accept=".docx" onChange={(e) => handleFileChange(e, 'template')} />
                {templateFile && <p className="file-name">{templateFile.name}</p>}
            </div>

            <div className="input-group">
                <label htmlFor="db-upload">2. Upload SQLite Database (.db)</label>
                <input id="db-upload" type="file" accept=".db, .sqlite, .sqlite3" onChange={(e) => handleFileChange(e, 'db')} />
                {dbFile && <p className="file-name">{dbFile.name}</p>}
            </div>

            <div className="input-group">
                <label htmlFor="report-date">3. Report Date (parm_month)</label>
                <input
                    id="report-date"
                    type="text"
                    value={reportDate}
                    onChange={(e) => setReportDate(e.target.value)}
                    placeholder="YYYY-MM"
                />
            </div>

            <div className="button-group">
                <button className="btn-secondary" onClick={handleDownloadSampleTemplate}>Download Sample</button>
                <button className="btn-secondary" onClick={handleHelpOpen}>Help</button>
                <button className="btn-secondary" onClick={handleSettingsOpen} disabled={!filesUploaded}>Settings</button>
                <button className="btn-primary" onClick={handleGenerate} disabled={!filesUploaded || isLoading}>
                    {isLoading ? 'Generating...' : 'Generate'}
                </button>
            </div>

            {isSettingsOpen && (
                <div className="modal-overlay">
                    <div className="modal-content">
                        <div className="modal-header">
                            <h2>Parameter Settings</h2>
                             <button className="close-button" onClick={handleSettingsClose}>&times;</button>
                        </div>
                        <table className="param-table">
                            <thead>
                                <tr>
                                    <th>Report Name</th>
                                    <th>Parameter Name</th>
                                    <th>SQL Query</th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                {params.map((param, index) => (
                                    <tr key={index}>
                                        <td><input type="text" value={param.report_name} onChange={e => handleParamChange(index, 'report_name', e.target.value)} /></td>
                                        <td><input type="text" value={param.parm_name} onChange={e => handleParamChange(index, 'parm_name', e.target.value)} /></td>
                                        <td><input type="text" value={param.SQL} onChange={e => handleParamChange(index, 'SQL', e.target.value)} /></td>
                                        <td>
                                            <button className="btn-danger" onClick={() => handleDeleteParam(index)}>Delete</button>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                        <div className="modal-footer">
                            <button className="btn-secondary" onClick={handleAddParam}>Add Parameter</button>
                            <button className="btn-primary" onClick={handleSaveChanges}>Save Changes</button>
                            <button className="btn-secondary" onClick={handleSettingsClose}>Cancel</button>
                        </div>
                    </div>
                </div>
            )}

            {isHelpOpen && (
                 <div className="modal-overlay">
                    <div className="modal-content">
                        <div className="modal-header">
                            <h2>How to Use Parameters in Your Word Template</h2>
                             <button className="close-button" onClick={handleHelpClose}>&times;</button>
                        </div>
                        <div className="help-content">
                            <p>To automatically insert data into your report, you need to add placeholders, or "parameters," to your Word document. The application will find these parameters and replace them with data from your SQLite database.</p>
                            
                            <h3>Parameter Syntax</h3>
                            <p>Parameters must be written in the following format: <code>[$parameter_name]</code></p>
                            <ul>
                                <li>They must start with <code>[$</code> and end with <code>]</code>.</li>
                                <li>The <code>parameter_name</code> inside must exactly match a <code>parm_name</code> you have defined in the Parameter Settings.</li>
                            </ul>

                            <h3>Special Built-in Parameter</h3>
                             <p><code>[$parm_month]</code>: This is a special parameter. It will be automatically replaced with the value you enter in the "Report Date" input field on the main page.</p>

                            <h3>Examples</h3>

                            <h4>1. Fetching a Single Value</h4>
                            <p>Imagine you have a parameter in your settings named <code>total_sales</code> with the SQL query <code>SELECT SUM(amount) FROM sales</code>.</p>
                            <p>In your Word document, you can write:</p>
                            <p><em>Total Sales for the Month: <code>[$total_sales]</code></em></p>
                            <p>The application will run the SQL query, and if the result is '54321', the final document will show:</p>
                            <p><em>Total Sales for the Month: 54321</em></p>

                            <h4>2. Embedding in a Sentence</h4>
                            <p>Let's say you have a parameter <code>top_product</code> with the SQL query <code>SELECT product_name FROM monthly_sales ORDER BY sales_volume DESC LIMIT 1</code>.</p>
                             <p>In your Word document, you can write:</p>
                            <p><em>Our best-selling item this month was the <code>[$top_product]</code>, showing strong market demand.</em></p>
                            <p>If the query returns 'SuperWidget', the output will be:</p>
                            <p><em>Our best-selling item this month was the SuperWidget, showing strong market demand.</em></p>


                            <h4>3. Using the Report Date</h4>
                            <p>You can use the built-in <code>[$parm_month]</code> parameter to dynamically add the report period to your document's title or headers.</p>
                            <p>In your Word document, you might have a title like:</p>
                            <p><em>Monthly Performance Review: <code>[$parm_month]</code></em></p>
                            <p>If you set the "Report Date" to '2025-09' in the app, the generated report will have the title:</p>
                            <p><em>Monthly Performance Review: 2025-09</em></p>
                        </div>
                         <div className="modal-footer">
                            <button className="btn-primary" onClick={handleHelpClose}>Close</button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

const container = document.getElementById('root');
const root = createRoot(container!);
root.render(<App />);
