import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import "./style.css";
import MediaList from "./MediaList";

export interface ICarouselProps {
  context: WebPartContext;
  styles: any;
  props: any;
}

export interface ICarousel {
  id: string;
  name: string;
  title: string;
  link:string;
  caption: string;
  type?: string;
  src: string;
  size: string;
  modified: string;
  seq?: number;
  seq2?: number;
}

export default function Carousel({
  context,
  styles,
  props,
}: ICarouselProps): JSX.Element {
  const sp = spfi(context.pageContext.web.absoluteUrl).using(SPFx(context));
  const [carousel, setCarousel] = useState<ICarousel[]>([]);
  const [sliderRef, setSliderRef] = useState(null);
  const [layout, setLayout] = useState<string>("");

  let height = props.height ? props.height / 16 : 500 / 16;
  let containerHeight = props.height ? parseInt(props.height) + 25 : 525;

  const handleAfterChange = (current: number) => {
    const item = carousel[current];

    if (item.type === "video") sliderRef.slickPause();
    else sliderRef.slickPlay();
  };

  useEffect(() => {
    if (carousel.length === 1) setLayout("1");
    else setLayout(props.layout);
  }, [props.layout]);

  const settings = {
    className: "",
    dots: true,
    infinite: true,
    slidesToShow: layout && layout !== "1" ? 3 : 1,
    slidesToScroll: 1, //props.layout === "1" ? 1 : 3,
    adaptiveHeight: false,
    autoplay: props.autoplay,
    speed: 1000,
    autoplaySpeed: props.autoplaySpeed,
    pauseOnHover: props.pauseOnHover,
    afterChange: handleAfterChange,
  };

  useEffect(() => {
    const getCorousel = async (): Promise<void> => {
      setCarousel([]);

      const files = await sp.web.lists
        .getByTitle(props.listName)
        .rootFolder.files.expand("ListItemAllFields")();

      const carouselItems: ICarousel[] = [];
      files.forEach((file) => {
        let filetype = "image";
        const name = file.Name.split(".");
        const fileSize =
          Number(file.Length) > 1000000
            ? Math.round((Number(file.Length) / 1000000) * 100) / 100 + "MB"
            : Math.round(Number(file.Length) / 1000) + "KB";

        if (
          name[1] === "png" ||
          name[1] === "jpeg" ||
          name[1] === "jpg" ||
          name[1] === "gif"
        )
          filetype = "image";
        else if (name[1] === "mp4") filetype = "video";
        else filetype = "audio";

        const el: ICarousel = {
          id: file.UniqueId,
          name: file.Name,
          title: file.Title,
          link: file["ListItemAllFields"].Link,
          caption: file["ListItemAllFields"].Caption,
          type: filetype,
          src: file.ServerRelativeUrl.replace(/#/g, "%23"),
          size: fileSize,
          modified: new Date(file.TimeLastModified).toLocaleString(),
          // seq: file["ListItemAllFields"].Sequence,
          // //seq2 is used to check if seq input field change
          // seq2: 1,
        };

        carouselItems.push(el);
      });

      if (files.length == 0) {
        /*const el: ICarousel = {
          id: "1",
          name: "Frame.png",
          title: "Frame.png",
          caption: "",
          type: "image",
          src: require("../image/Frame.png"),
          size: "",
          modified: "",
          // seq: 1,
          // //seq2 is used to check if seq input field change
          // seq2: 1,
        };

        carouselItems.push(el);*/
      }

      if (carouselItems.length === 1) {
        settings.slidesToShow = 1;
        setLayout("1");
      } else {
        settings.slidesToShow = layout && layout !== "1" ? 3 : 1;
        setLayout(props.layout);
      }

      carouselItems.sort((a, b) => a.seq - b.seq);
      setCarousel(carouselItems);
    };

    //if (props.listName !== undefined && carousel.length === 0)
    void getCorousel();
  }, [props.listName]);

  return (
    <>
      <div
        className="content"
        style={{
          borderRadius: props.borderRadius + "px",
          minHeight: containerHeight + "px",
          backgroundColor: props.backgroundColor,
        }}
      >
        <Slider ref={setSliderRef} {...settings}>
          {carousel.map((card, index) => (
            <>
              <div style={{textAlign:'center', fontSize:'18px', fontWeight:'bold', margin:'10px 5px'}}>
                <a href={card.link}>{card.title}</a>
              </div>
              <div
                key={card.id}
                className={
                  layout !== "2"
                    ? "carousel-item"
                    : "carousel-item carousel-item-padd"
                }
              >
                {card.type === "image" && (
                  <img
                    src={card.src}
                    alt={card.title}
                    style={{
                      objectFit: layout === "1" ? props.displayStyle : "fill",
                      borderRadius: props.borderRadius + "px",
                      height: height + "rem",
                    }}
                  />
                )}
                {card.type === "video" && (
                  <>
                    <video
                      controls
                      autoPlay
                      //muted
                      loop
                      style={{
                        borderRadius: props.borderRadius + "px",
                        height: height + "rem",
                      }}
                    >
                      <source src={card.src} type="video/mp4"></source>
                    </video>
                  </>
                )}
                {card.type === "audio" && (
                  <>
                    <audio
                      controls
                      autoPlay
                      //muted
                      loop
                      style={{
                        borderRadius: props.borderRadius + "px",
                        height: height + "rem",
                        width:'100%'
                      }}
                    >
                      <source src={card.src} type="audio/mp4"></source>
                      Your browser does not support the audio element.
                    </audio>
                  </>
                )}

                <div className={"caption caption-" + props.captionPosition}>
                  {/*card.title && "Caption" */}
                </div>
                <div
                  className={"text text-" + props.captionPosition}
                  style={{
                    fontSize: props.captionFontSize + "px",
                    fontWeight: props.captionWeight,
                    color: props.captionColor,
                  }}
                >
                  {card.caption}
                </div>
              </div>
            </>
            
          ))}
        </Slider>
        <a className="prev" onClick={sliderRef?.slickPrev}>
          &#10094;
        </a>
        <a className="next" onClick={sliderRef?.slickNext}>
          &#10095;
        </a>
      </div>

      {/* <div className="content">
        <div className="controls">
          <button onClick={sliderRef?.slickPrev}>
            <FaChevronLeft />
          </button>
          <button onClick={sliderRef?.slickNext}>
            <FaChevronRight />
          </button>
        </div>
        <Slider ref={setSliderRef} {...settings}>
          {hotelCards.map((card, index) => (
            <div key={index} className="card">
              <img
                src={card.imageSrc}
                alt={card.title}
                className="card-image"
              />
              <div className="text-info">
                <div className="card-header">
                  <h2>{card.title}</h2>
                  <span>{card.pricingText}</span>
                </div>
                <p>{card.description}</p>
                <ul>
                  {card.features.map((feature, index) => (
                    <li key={index}>{feature}</li>
                  ))}
                </ul>
              </div>
            </div>
          ))}
        </Slider>
      </div> */}
    </>
  );
}
