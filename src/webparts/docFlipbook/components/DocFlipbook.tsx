import * as React from 'react';
import styles from './DocFlipbook.module.scss';
import "bootstrap/dist/css/bootstrap.min.css";

import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import FlipBookViewer from './FlipBookViewer';

import "bootstrap-icons/font/bootstrap-icons.css";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IProps {
  context: WebPartContext;
  libraryName: string;
}
interface IDocumentItem {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  Year: string;
  UniqueId: string;
}
const DocFlipbook: React.FC<IProps> = ({ context, libraryName }) => {

 const currentYear = new Date().getFullYear().toString();

// const [docs, setDocs] = React.useState<any[]>([]);
// const [filteredDocs, setFilteredDocs] = React.useState<any[]>([]);
const [docs, setDocs] = React.useState<IDocumentItem[]>([]);
const [filteredDocs, setFilteredDocs] = React.useState<IDocumentItem[]>([]);
const [selectedYear, setSelectedYear] = React.useState<string>(currentYear);
const [selectedFile, setSelectedFile] = React.useState<string | null>(null);

const years = Array.from({ length: 4 }, (_, i) =>
  (new Date().getFullYear() - i).toString()
);

  // ✅ Create SP instance (PnP v3)
  const sp: SPFI = React.useMemo(() => {
    return spfi().using(SPFx(context));
  }, [context]);



 const fetchDocuments = async (): Promise<void> => {
  try {
    const items = await sp.web.lists
      .getByTitle("DigitalFlipbook")
      .items
      .select("Id", "FileLeafRef", "FileRef", "Year")
      .filter("FSObjType eq 0 and substringof('.pdf', FileLeafRef)")();

    setDocs(items);

  const currentYearDocs = items.filter(
  item => item.Year === currentYear
);
    setFilteredDocs(currentYearDocs);

  } catch (error) {
    console.error("Error fetching documents:", error);
  }
};

React.useEffect(() => {
  fetchDocuments().catch((error) => {
    console.error("Error fetching documents:", error);
  });
}, []);
  
  const handleYearChange = (year: string): void => {
  setSelectedYear(year);

  const filtered = docs.filter(
    (item: IDocumentItem) => item.Year === year
  );

  setFilteredDocs(filtered);
};

  return (
    // <div className={styles.container}>
    <div className={styles.commonSectionstyle}>
     {!selectedFile && (
      <div className={styles.pageContainer}>
        <div className={styles.banner}>
          <div className={styles.bannerIcon}>
            <i className="bi bi-journal-bookmark-fill" />
          </div>

          <div>
            <h3>Digital Library Visualization</h3>

            <p>
              View and access month-wise PDF documents organized by the selected
              year.
            </p>
          </div>
        </div>
      </div>
    )}
      <div className="d-flex flex-column flex-md-row align-items-md-center justify-content-between gap-3 mb-4 mt-3">
       {!selectedFile && (
          <h5 className="mb-0 text-nowrap">
            {filteredDocs.length} {filteredDocs.length === 1 ? "File" : "Files"} Found
          </h5>
        )}

        {!selectedFile && (
            <div className="d-flex flex-column flex-sm-row align-items-sm-center ms-md-auto w-100 w-md-auto justify-content-md-end">
              <label className="me-sm-2 mb-2 mb-sm-0">Year :</label>

              <select
                className="form-select"
                style={{ maxWidth: "200px", width: "100%" }}
                value={selectedYear}
                onChange={(e) => handleYearChange(e.target.value)}
              >
                {years.map((year) => (
                  <option key={year} value={year}>
                    {year}
                  </option>
                ))}
              </select>
            </div>
          )}
      </div>
    {!selectedFile && (
  <>
    {filteredDocs.length > 0 ? (
      <div className={styles.cardGrid}>
        {filteredDocs.map((doc, index) => {

          const thumbnailUrl = `${context.pageContext.web.absoluteUrl}/_layouts/15/getpreview.ashx?path=${doc.FileRef}`;

          return (
            <div
              key={index}
              className={styles.card}
              onClick={() => setSelectedFile(`${window.location.origin}${doc.FileRef}`)}
            >
              <img src={thumbnailUrl} className={styles.thumbnail} />
              <div className={styles.title}>{doc.FileLeafRef}</div>
            </div>
          );
        })}
      </div>
    ) : (
      <div className={styles.noDataMessage}>
        No documents found for the selected year.
      </div>
    )}
  </>
)}
      {selectedFile && (
        <FlipBookViewer
          fileUrl={`${selectedFile}`}
          onClose={() => setSelectedFile(null)}
        />
      )}

    </div>
  );
};

export default DocFlipbook;