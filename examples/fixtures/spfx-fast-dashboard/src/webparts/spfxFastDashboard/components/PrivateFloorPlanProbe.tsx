import * as React from "react";
import { Button, MessageBar, MessageBarBody, Text } from "@fluentui/react-components";
import { DriveFloorPlanMap } from "@spfx-fast/core/react";
import type { createClient } from "@spfx-fast/core";

type Client = ReturnType<typeof createClient>;

async function createSyntheticFloorPlan(): Promise<Blob> {
  const canvas = document.createElement("canvas");
  canvas.width = 800;
  canvas.height = 600;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas is unavailable.");
  context.fillStyle = "#f4f8f7";
  context.fillRect(0, 0, 800, 600);
  context.fillStyle = "#fff";
  context.fillRect(24, 24, 752, 552);
  const rooms = [
    { x: 40, y: 40, width: 300, height: 208, color: "#eaf4f8" },
    { x: 360, y: 40, width: 400, height: 208, color: "#fff4d9" },
    { x: 40, y: 288, width: 300, height: 272, color: "#e8f4ed" },
    { x: 360, y: 288, width: 400, height: 272, color: "#edf0fa" },
  ];
  for (const room of rooms) {
    context.fillStyle = room.color;
    context.fillRect(room.x, room.y, room.width, room.height);
  }
  context.strokeStyle = "#66818a";
  context.lineWidth = 8;
  context.strokeRect(24, 24, 752, 552);
  context.lineWidth = 6;
  context.beginPath();
  context.moveTo(350, 40); context.lineTo(350, 560); context.moveTo(40, 268); context.lineTo(760, 268);
  context.stroke();
  context.fillStyle = "#294c55";
  context.font = "600 21px sans-serif";
  context.fillText("RECEPTION", 64, 78);
  context.fillText("MEETING ROOM", 392, 78);
  context.fillText("OPEN WORKSPACE", 64, 326);
  context.fillText("QUIET ROOMS", 392, 326);
  context.fillStyle = "#fff";
  context.strokeStyle = "#9bb0b4";
  context.lineWidth = 3;
  context.fillRect(78, 108, 194, 48); context.strokeRect(78, 108, 194, 48);
  context.fillRect(450, 106, 220, 78); context.strokeRect(450, 106, 220, 78);
  context.fillRect(80, 362, 88, 64); context.strokeRect(80, 362, 88, 64);
  context.fillRect(194, 362, 88, 64); context.strokeRect(194, 362, 88, 64);
  context.fillRect(432, 368, 118, 112); context.strokeRect(432, 368, 118, 112);
  context.fillRect(574, 368, 142, 112); context.strokeRect(574, 368, 142, 112);
  context.fillStyle = "#58757d";
  context.font = "600 14px sans-serif";
  context.fillText("FRONT DESK", 116, 190);
  context.fillText("6 SEATS · DISPLAY", 482, 210);
  context.fillText("DESK BANK A", 104, 454);
  context.fillText("QUIET ZONE · 4 ROOMS", 474, 516);
  return new Promise((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("Could not create the sample image.")), "image/png"));
}

/** Live check of authenticated private-image reads using one disposable synthetic image. */
export function PrivateFloorPlanProbe({ client, driveId }: { client: Client; driveId: string }): React.ReactElement {
  const [file, setFile] = React.useState<{ id: string; name: string }>();
  const [busy, setBusy] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [error, setError] = React.useState(false);

  async function createAndRead(): Promise<void> {
    if (busy || !window.confirm("Upload one synthetic PNG to the library, read it through the signed-in Graph client, then remove it?")) return;
    setBusy(true);
    setError(false);
    setStatus("");
    try {
      const name = `SPFxFastFloorPlan_${crypto.randomUUID()}.png`;
      const uploaded = await client.uploadDriveItem(driveId, "root", name, await createSyntheticFloorPlan());
      if (typeof uploaded.id !== "string" || !uploaded.id) throw new Error("Graph did not return the uploaded image ID.");
      setFile({ id: uploaded.id, name });
      setStatus("Image uploaded. The map is now reading its private bytes through the cached Graph hook.");
    } catch {
      setError(true);
      setStatus("The test image could not be uploaded. Check Graph file permission and library access.");
    } finally {
      setBusy(false);
    }
  }

  async function removeImage(): Promise<void> {
    if (!file || busy) return;
    setBusy(true);
    setError(false);
    try {
      await client.deleteDriveItem(driveId, file.id);
      setFile(undefined);
      setStatus("Temporary image removed.");
    } catch {
      setError(true);
      setStatus(`Image cleanup failed. Remove ${file.name} from the library manually.`);
    } finally {
      setBusy(false);
    }
  }

  return <section aria-label="Private floor-plan image check" style={{ display: "grid", gap: 12 }}>
    <h2>Private floor-plan image</h2>
    <Text>Uploads a synthetic plan, reads its bytes through the signed-in Graph client and removes the temporary file.</Text>
    {!file ? <Button appearance="secondary" disabled={busy} onClick={createAndRead}>{busy ? "Uploading image…" : "Test private floor-plan image"}</Button> : <>
      <Text>Temporary image: {file.name}</Text>
      <DriveFloorPlanMap
        client={client}
        driveId={driveId}
        itemId={file.id}
        title="Synthetic facilities"
        locations={[
          { id: "reception", label: "Reception desk", x: 18, y: 22, detail: "Front desk" },
          { id: "meeting-1", label: "Meeting 1", x: 68, y: 25, detail: "6 seats" },
          { id: "desk-1", label: "Desk 1", x: 26, y: 64, detail: "Available" },
          { id: "quiet-room", label: "Quiet rooms", x: 80, y: 70, detail: "Four phone rooms" },
        ]}
      />
      <Button appearance="secondary" disabled={busy} onClick={removeImage}>{busy ? "Removing image…" : "Remove temporary floor-plan image"}</Button>
    </>}
    {status ? error ? <MessageBar intent="error"><MessageBarBody role="alert">{status}</MessageBarBody></MessageBar> : <p role="status" aria-live="polite">{status}</p> : null}
  </section>;
}
