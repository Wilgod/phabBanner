import * as fontawesome from "@fortawesome/free-solid-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/files";
import { IFileAddResult } from "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import arraySort from "array-sort";
import * as React from "react";
import { useEffect, useState } from "react";
import { ICarousel } from "./Carousel";
import styles from "./MediaList.module.scss";

export interface IMediaListProps {
  context: WebPartContext;
  props: any;
}

export default function MediaList({ context, props }: IMediaListProps) {
  const [carousel, setCarousel] = useState<any[]>([]);
  const [uploadFile, setUploadFile] = useState<File>();
  const [sequence, setSequence] = useState(Number);
  const sp = spfi(context.pageContext.web.absoluteUrl).using(SPFx(context));

  const getCarousel = async () => {
    // let url = context.pageContext.web.absoluteUrl
    const result = await sp.web
      .getFolderByServerRelativePath(props.listName)
      .files.expand(
        "ListItemAllFields",
        "ListItemAllFields/FileRef",
        "ListItemAllFields/EncodedAbsThumbnailUrl"
      )();
    let slides: ICarousel[] = [];
    result.map((item) =>
      slides.push({
        id: item.UniqueId,
        name: item.Name,
        title: item.Title,
        caption: item["ListItemAllFields"].Caption,
        src: item.ServerRelativeUrl.replace(/#/g, "%23"),
        size:
          Number(item.Length) > 1000000
            ? Math.round((Number(item.Length) / 1000000) * 100) / 100 + "MB"
            : Math.round(Number(item.Length) / 1000) + "KB",
        modified: new Date(item.TimeLastModified).toLocaleString(),
        seq: item["ListItemAllFields"].Sequence,
        seq2: item["ListItemAllFields"].Sequence,
      })
    );
    arraySort(slides, "seq");
    slides.map(async (item, i) => {
      if (++i != item.seq) {
        item.seq = i;
        item.seq2 = i;
        let slide = await sp.web.getFileById(item.id).getItem();
        await slide.update({ Sequence: i });
      }
    });
    setCarousel(slides);
  };
  function addAfter(array, index, newItem) {
    return [...array.slice(0, index), newItem, ...array.slice(index)];
  }

  const _isInt = (value) => {
    return (
      !isNaN(value) &&
      (function (x) {
        return (x | 0) === x;
      })(parseFloat(value))
    );
  };

  const handleUpload = async () => {
    if (!uploadFile || !_isInt(sequence) || sequence == 0) {
      alert("Sequence must be an integer and an image file must be specified.");
      return;
    }
    const nameArr = carousel.map((item) => item.name);
    if (nameArr.includes(uploadFile.name)) {
      alert(`${uploadFile.name} already exists`);
      return;
    }
    if (confirm("Confirm Upload?")) {
      const fileNamePath = uploadFile.name; //encodeURI(uploadFile.name);
      let result: IFileAddResult;
      // you can adjust this number to control what size files are uploaded in chunks
      if (uploadFile.size <= 10485760) {
        // small upload
        result = await sp.web
          .getFolderByServerRelativePath(props.listName)
          .files.addUsingPath(fileNamePath, uploadFile);
      } else {
        // large upload
        result = await sp.web
          .getFolderByServerRelativePath(props.listName)
          .files.addChunked(
            fileNamePath,
            uploadFile,
            (data) => {
              console.log(
                Math.round((data.currentPointer / data.fileSize) * 100) + "%"
              );
            },
            true
          );
      }

      const slide = {
        id: result.data.UniqueId,
        name: result.data.Name,
        src: result.data.ServerRelativeUrl.replace(/#/g, "%23"),
        size:
          Number(result.data.Length) > 1000000
            ? Math.round((Number(result.data.Length) / 1000000) * 100) / 100 +
              "MB"
            : Math.round(Number(result.data.Length) / 1000) + "KB",
        modified: new Date(result.data.TimeLastModified).toLocaleString(),
        seq: [sequence, sequence],
      };

      let newSlides = addAfter(carousel, sequence - 1, slide);
      newSlides.map(async (item, i) => {
        item.seq = ++i;
        item.seq2 = i;
        let slide = await sp.web.getFileById(item.id).getItem();
        await slide.update({ Sequence: i });
      });
      setCarousel(newSlides);
    }
  };
  const handleDelete = async (fileName) => {
    if (confirm("Confirm Delete?")) {
      await sp.web
        .getFolderByServerRelativePath(props.listName)
        .files.getByUrl(fileName)
        .delete();

      void getCarousel();
      alert(`${fileName} has been removed.`);
    }
  };

  const handleInputChange = (e) => {
    const value = e.target.value;
    const index = e.target.dataset.index;
    if (!_isInt(value) || value == 0) {
      console.log("Sequence must be an integer greater than 0.");
    }
    setCarousel(
      carousel.map((item, i) =>
        i !== parseInt(index) ? item : { ...item, seq2: parseInt(value) }
      )
    );
  };
  const handleUpdate = (index: number, seq: number) => {
    if (confirm("Confirm Update?")) {
      const prevSeq = carousel[index].seq;
      //debugger
      let newSlides = [...carousel];
      const insertItem = newSlides.splice(prevSeq - 1, 1)[0];
      newSlides = addAfter(newSlides, seq - 1, insertItem);
      newSlides = newSlides.map((item, i) => ({
        ...item,
        seq: i + 1,
        seq2: i + 1,
      }));
      //debugger
      newSlides.map(async (item, i) => {
        let slide = await sp.web.getFileById(item.id).getItem();
        await slide.update({ Sequence: ++i });
      });
      setCarousel(newSlides);
    }
  };

  useEffect(() => {
    void getCarousel();
  }, []);

  return (
    <>
      <div className={styles.lbCarouselMediaList}>
        {!props.hideUpload && (
          <section className={styles.uploadContainer}>
            <h6>{props.listName}</h6>
            <div className={styles.fileInputWrap}>
              <div>
                {props.hideUpload} Sequence#{" "}
                <input
                  type="number"
                  min="1"
                  onChange={(e) => setSequence(Number(e.target.value))}
                />
              </div>
              <div>
                <input
                  type="file"
                  onChange={(e) => {
                    if (e.target.files === null) return;
                    setUploadFile(e.target.files[0]);
                  }}
                  accept=".jpg,.jpeg,.png,.webp,.gif,.mp4"
                />
              </div>
              <div>
                <button onClick={handleUpload}>Submit</button>
              </div>
            </div>
          </section>
        )}
        <section className={styles.listing}>
          <table>
            <tr>
              <th className={styles.name}>Name</th>
              <th className={styles.seqence}>Sequence</th>
              <th className={styles.size}>Size</th>
              <th className={styles.modified}>Modified</th>
              <th className={styles.delete}>Delete?</th>
            </tr>
            {carousel.map((item, i) => (
              <tr>
                <td className={styles.name}>
                  <span className={styles.thumbnail}>
                    <img src={item.src} width={50} />
                  </span>
                  <span className={styles.imgName}>{item.name}</span>
                </td>
                <td className={styles.seqence}>
                  <div className={`${styles.colContent} ${styles.seqContent}`}>
                    <input
                      type="number"
                      value={item.seq2}
                      min="1"
                      name="sequence"
                      onChange={handleInputChange}
                      data-index={i}
                    />
                    <FontAwesomeIcon
                      onClick={() => handleUpdate(i, item.seq2)}
                      title="Update Sequence"
                      className={`${styles.pointerCursor} ${styles.edit} ${
                        item.seq === item.seq2 && styles.hidden
                      }`}
                      icon={fontawesome["faPen"]}
                      color="#444"
                      size="1x"
                    />
                  </div>
                </td>
                <td className={styles.size}>
                  <div className={styles.colContent}>{item.size}</div>
                </td>
                <td className={styles.modified}>
                  <div className={styles.colContent}>{item.modified}</div>
                </td>
                <td className={styles.delete}>
                  <div className={styles.colContent}>
                    <FontAwesomeIcon
                      className={`${styles.pointerCursor} ${styles.delete}`}
                      onClick={() => handleDelete(item.name)}
                      icon={fontawesome["faTrash"]}
                      color="#993333"
                      size="1x"
                    />
                  </div>
                </td>
              </tr>
            ))}
          </table>
        </section>
      </div>
    </>
  );
}
