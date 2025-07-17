import { Riple } from "react-loading-indicators";

export default function LoadingAnimation() {
  return (
    // <svg xmlns="http://www.w3.org/2000/svg" style={{ margin: "auto", background: "none", display: "block", shapeRendering: "auto" }} width="40px" height="40px" viewBox="0 0 100 100" preserveAspectRatio="xMidYMid">
    //     <circle cx="31" cy="50" fill="#00964b" r="19">
    //         <animate attributeName="cx" repeatCount="indefinite" dur="1.5384615384615383s" keyTimes="0;0.5;1" values="31;69;31" begin="-0.7692307692307692s"></animate>
    //     </circle>
    //     <circle cx="69" cy="50" fill="#ffd100" r="19">
    //         <animate attributeName="cx" repeatCount="indefinite" dur="1.5384615384615383s" keyTimes="0;0.5;1" values="31;69;31" begin="0s"></animate>
    //     </circle>
    //     <circle cx="31" cy="50" fill="#00964b" r="19">
    //         <animate attributeName="cx" repeatCount="indefinite" dur="1.5384615384615383s" keyTimes="0;0.5;1" values="31;69;31" begin="-0.7692307692307692s"></animate>
    //         <animate attributeName="fill-opacity" values="0;0;1;1" calcMode="discrete" keyTimes="0;0.499;0.5;1" dur="1.5384615384615383s" repeatCount="indefinite"></animate>
    //     </circle>
    // </svg>
    <div className="flex justify-center items-center">
      <Riple color="#8dbd3f" size="small" text="" textColor="" />
    </div>
  );
}
