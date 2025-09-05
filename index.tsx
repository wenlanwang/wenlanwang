
import React, { useState, useEffect, useRef } from 'react';
import { createRoot } from 'react-dom/client';

declare const initSqlJs: any;
declare const PizZip: any;
declare const docxtemplater: any;
declare const saveAs: any;

interface Param {
    report_name: string;
    parm_name: string;
    SQL: string;
    report_date: string;
}

const App: React.FC = () => {
    const [db, setDb] = useState<any>(null);
    const [templateFile, setTemplateFile] = useState<File | null>(null);
    const [templateFileName, setTemplateFileName] = useState<string>('');
    const [dbFileName, setDbFileName] = useState<string>('');
    const [reportDate, setReportDate] = useState<string>('');
    const [error, setError] = useState<string>('');
    const [generatedFile, setGeneratedFile] = useState<Blob | null>(null);
    const [isSettingsOpen, setIsSettingsOpen] = useState<boolean>(false);
    const [isHelpOpen, setIsHelpOpen] = useState<boolean>(false);
    const [params, setParams] = useState<Param[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(false);

    const dbWorkerRef = useRef<Worker | null>(null);

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
            setDbFileName(file.name);
            const buffer = await file.arrayBuffer();
            try {
                const SQL = await initSqlJs({ locateFile: (file: string) => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${file}` });
                const database = new SQL.Database(new Uint8Array(buffer));
                setDb(database);
                setError('');
            } catch (err) {
                console.error(err);
                setError('Failed to load database. Please ensure it is a valid SQLite file.');
            }
        } else if (fileType === 'template') {
            setTemplateFile(file);
            setTemplateFileName(file.name);
        }
    };

    const fetchParams = () => {
        if (!db) {
            setError('Database not loaded.');
            return;
        }
        try {
            const res = db.exec("SELECT report_name, parm_name, SQL, report_date FROM parm");
            if (res.length > 0) {
                const data = res[0].values.map((row: any) => ({
                    report_name: row[0],
                    parm_name: row[1],
                    SQL: row[2],
                    report_date: row[3],
                }));
                setParams(data);
            }
        } catch (err) {
            console.error(err);
            setError('Failed to fetch parameters from the "parm" table. Make sure the table exists and has the correct schema.');
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

    const handleSaveChanges = () => {
        if (!db) return;
        try {
            db.exec("BEGIN TRANSACTION;");
            db.exec("DELETE FROM parm;");
            params.forEach(p => {
                db.run("INSERT INTO parm (report_name, parm_name, SQL, report_date) VALUES (?, ?, ?, ?)", [p.report_name, p.parm_name, p.SQL, p.report_date]);
            });
            db.exec("COMMIT;");
            handleSettingsClose();
        } catch (err) {
            db.exec("ROLLBACK;");
            console.error(err);
            setError('Failed to save changes to the database.');
        }
    };


    const handleGenerate = async () => {
        if (!templateFile || !db) {
            setError('Please upload both a template file and a database file.');
            return;
        }

        setError('');
        setIsLoading(true);
        setGeneratedFile(null);

        if (typeof PizZip === 'undefined' || typeof docxtemplater === 'undefined') {
            setError('A required library (PizZip or docxtemplater) could not be loaded. Please check your internet connection, disable any ad-blockers that might interfere with script loading, and refresh the page.');
            setIsLoading(false);
            return;
        }

        try {
            const templateBuffer = await templateFile.arrayBuffer();
            const zip = new PizZip(templateBuffer);

            const nullGetter = (part: any) => {
                // For unresolved placeholders like [$param], return the original text to avoid errors.
                if (part.type === "placeholder") {
                    return `[$${part.value}]`;
                }
                return "";
            };

            const doc = new docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                delimiters: {
                    start: '[$',
                    end: ']',
                },
                nullGetter: nullGetter,
            });

            const content = doc.getFullText() as string;
            const placeholders = [...new Set(content.match(/\[\$(.*?)\]/g) || [])];
            const data: { [key: string]: any } = { parm_month: reportDate };

            for (const placeholder of placeholders) {
                const parm_name = placeholder.substring(2, placeholder.length - 1);
                if (parm_name === 'parm_month') continue;

                try {
                    // Use prepared statements to prevent SQL injection.
                    const stmt = db.prepare("SELECT SQL FROM parm WHERE parm_name = :parm_name");
                    stmt.bind({ ':parm_name': parm_name });
                    if (stmt.step()) {
                        const sqlQuery = stmt.get()[0] as string;
                        // Execute the user-provided SQL
                        const queryResult = db.exec(sqlQuery);
                        if (queryResult.length > 0 && queryResult[0].values.length > 0) {
                            // Use the first cell of the first row as the result
                            data[parm_name] = queryResult[0].values[0][0];
                        } else {
                            data[parm_name] = `[No result for ${parm_name}]`;
                        }
                    }
                    // If param is not in DB, stmt.step() is false. nullGetter will handle it.
                    stmt.free();
                } catch (e: any) {
                    console.error(`Error executing SQL for ${parm_name}:`, e);
                    data[parm_name] = `[SQL Error for ${parm_name}]`;
                }
            }

            doc.render(data);

            const out = doc.getZip().generate({
                type: "blob",
                mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });

            setGeneratedFile(out);
        } catch (err: any) {
            console.error(err);
            let errorMessage = 'An error occurred during report generation. Please check the template format and parameter names.';
            if (err.properties && err.properties.errors) {
                const specificError = err.properties.errors[0].properties.explanation;
                errorMessage += ` Details: ${specificError}`;
            } else if (err.message) {
                 errorMessage += ` Details: ${err.message}`;
            }
            setError(errorMessage);
        } finally {
            setIsLoading(false);
        }
    };

    const handleDownload = () => {
        if (!generatedFile || !templateFile) return;
        const originalName = templateFile.name.replace(/\.docx$/, '');
        const newFileName = `${originalName}_${reportDate}.docx`;
        saveAs(generatedFile, newFileName);
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
                <label htmlFor="template-upload">1. 请上传word文档报告模板 (.docx)</label>
                <input id="template-upload" type="file" accept=".docx" onChange={(e) => handleFileChange(e, 'template')} />
                {templateFileName && <p className="file-name">{templateFileName}</p>}
            </div>

            <div className="input-group">
                <label htmlFor="db-upload">2. 请上传SQLite数据库文件 (.db)</label>
                <input id="db-upload" type="file" accept=".db, .sqlite, .sqlite3" onChange={(e) => handleFileChange(e, 'db')} />
                 {dbFileName && <p className="file-name">{dbFileName}</p>}
            </div>

            <div className="input-group">
                <label htmlFor="report-date">3. 报告月份 (parm_month)</label>
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
                <button className="btn-secondary" onClick={handleSettingsOpen} disabled={!db}>Settings</button>
                <button className="btn-primary" onClick={handleGenerate} disabled={!templateFile || !db || isLoading}>
                    {isLoading ? 'Generating...' : 'Generate'}
                </button>
                {generatedFile && (
                    <button className="btn-success" onClick={handleDownload}>Download Report</button>
                )}
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
                            {/* FIX: Corrected malformed HTML tag. Was `code>...` and is now `<code>...</code>` */}
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
