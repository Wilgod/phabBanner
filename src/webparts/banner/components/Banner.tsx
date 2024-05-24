import * as React from "react";
import type { IBannerProps } from "./IBannerProps";
import styles from "./Banner.module.scss";
import Carousel from "./Carousel";
const workbenchId = document.getElementById("workbenchPageContent");
/*const canvasZone = document.querySelector(".CanvasZone");
if (workbenchId !== null) {
  workbenchId.style.maxWidth = "none";
}
/*if (canvasZone !== null) {
  (canvasZone as HTMLElement).style.maxWidth =
    "none";
  (
    canvasZone.children[0] as HTMLElement
  ).style.maxWidth = "none";
}*/
/*if (document.querySelector(".ms-compositeHeader") !== null) {
  (document.querySelector(".ms-compositeHeader") as HTMLElement).style.display =
    "none";
}*/

export default class Banner extends React.Component<IBannerProps, {}> {
  public render(): React.ReactElement<IBannerProps> {
    const {
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.sbCarousel} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <Carousel
          context={this.props.context}
          styles={styles}
          props={this.props}
        />
      </section>
    );
  }
}
