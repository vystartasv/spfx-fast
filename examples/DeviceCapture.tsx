import * as React from "react";
import { Button, Field, Input, MessageBar, MessageBarBody } from "@fluentui/react-components";

type BarcodeDetectorLike = {
  detect(source: ImageBitmap): Promise<Array<{ rawValue?: string }>>;
};

/** Native camera/photo, barcode, and location inputs; all permissions are requested after user action. */
export function DeviceCapture({ onLocation, onBarcode }: {
  onLocation: (point: { latitude: number; longitude: number; accuracy: number }) => void;
  onBarcode: (value: string) => void;
}) {
  const [barcode, setBarcode] = React.useState("");
  const [message, setMessage] = React.useState("");
  const [busy, setBusy] = React.useState(false);

  async function scanPhoto(file?: File) {
    if (!file) return;
    const Detector = (globalThis as typeof globalThis & { BarcodeDetector?: new (options?: { formats?: string[] }) => BarcodeDetectorLike }).BarcodeDetector;
    if (!Detector || typeof createImageBitmap !== "function") {
      setMessage("Barcode scanning is unavailable in this browser. Enter the code below.");
      return;
    }
    setBusy(true);
    setMessage("");
    try {
      const bitmap = await createImageBitmap(file);
      try {
        const found = await new Detector().detect(bitmap);
        const value = found[0]?.rawValue?.trim();
        if (!value) throw new Error("No barcode found in that photo.");
        setBarcode(value);
        onBarcode(value);
      } finally {
        bitmap.close();
      }
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Could not scan this photo.");
    } finally {
      setBusy(false);
    }
  }

  function captureLocation() {
    if (!navigator.geolocation) {
      setMessage("Location is unavailable in this browser.");
      return;
    }
    setBusy(true);
    navigator.geolocation.getCurrentPosition(
      ({ coords }) => {
        try {
          onLocation({ latitude: coords.latitude, longitude: coords.longitude, accuracy: coords.accuracy });
          setMessage(`Location captured (accuracy about ${Math.round(coords.accuracy)} m).`);
        } catch {
          setMessage("Location was read but could not be saved by the app.");
        } finally {
          setBusy(false);
        }
      },
      (error) => {
        setMessage(error.message || "Location could not be read.");
        setBusy(false);
      },
      { enableHighAccuracy: true, timeout: 10_000, maximumAge: 30_000 },
    );
  }

  return <section aria-label="Device capture" style={{ display: "grid", gap: 12 }}>
    <Field label="Take a photo to scan a barcode">
      <input aria-label="Take or select a barcode photo" type="file" accept="image/*" capture="environment" disabled={busy} onChange={(event) => void scanPhoto(event.currentTarget.files?.[0])} />
    </Field>
    <Field label="Barcode" hint="Enter it manually when photo scanning is unavailable.">
      <Input value={barcode} onChange={(_, data) => setBarcode(data.value)} onBlur={() => { if (barcode.trim()) onBarcode(barcode.trim()); }} />
    </Field>
    <Button type="button" disabled={busy} onClick={captureLocation}>{busy ? "Working…" : "Capture current location"}</Button>
    {message ? <MessageBar><MessageBarBody>{message}</MessageBarBody></MessageBar> : null}
  </section>;
}
