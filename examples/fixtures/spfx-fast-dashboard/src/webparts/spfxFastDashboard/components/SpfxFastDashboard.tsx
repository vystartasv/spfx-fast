import * as React from "react";
import type { ISpfxFastDashboardProps } from "./ISpfxFastDashboardProps";
import { WorkDashboard } from "./WorkDashboard";

export default function SpfxFastDashboard(props: ISpfxFastDashboardProps): React.ReactElement {
  return <WorkDashboard {...props} />;
}
