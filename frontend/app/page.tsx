"use client";

import React, { useMemo, useState } from "react";
import JSZip from "jszip";
import * as XLSX from "xlsx";

type SlideTypeId = "org_change" | "new_tools" | "cw_risk_assessment";

type SlideType = {
  id: SlideTypeId;
  label: string;
  description?: string;
};

type CellSpec =
  | { type: "cell"; ref: string }
  | { type: "join"; refs: string[]; joinWith: string }
  | { type: "const"; value: string }
  | {
      type: "month_year";
      ref: string;
      format?: "Mon YYYY" | "MMMM YYYY" | "MM/YYYY";
    };

type SlideMapping = Record<number, Record<string, CellSpec>>;

/**
 * -----------------------------
 * MAPPINGS
 * -----------------------------
 */

const ORG_CHANGE_MAPPING: SlideMapping = {
  1: {
    "NAME OF PROJECT": { type: "cell", ref: "F2" },
    "TYPE OF PROJECT": { type: "cell", ref: "K2" },
  },
  3: {
    "[Description]": { type: "cell", ref: "M2" },
  },
  4: {
    "[L2/L3]": { type: "cell", ref: "I2" },
    "[Owner]": { type: "cell", ref: "G2" },
    "[Lead]": { type: "cell", ref: "H2" },
    "[Comms]": { type: "cell", ref: "J2" },
  },
  5: {
    "[Date]": { type: "cell", ref: "N2" },
    "[Phases]": { type: "cell", ref: "P2" },
  },
  6: {
    "[1]": { type: "cell", ref: "Q2" },
    "[2]": { type: "cell", ref: "R2" },
    "[3]": { type: "join", refs: ["S2", "T2"], joinWith: " " },
    "[4]": { type: "cell", ref: "V2" },
  },
  7: {
    "[1]": { type: "cell", ref: "W2" },
  },
  8: {
    "[1]": { type: "cell", ref: "L2" },
    "[2]": { type: "join", refs: ["AA2", "AB2"], joinWith: " " },
    "[3]": { type: "cell", ref: "Y2" },
  },
  9: {
    "[1]": { type: "cell", ref: "Z2" },
    "[2]": { type: "cell", ref: "AD2" },
  },
  10: {
    "[1]": { type: "cell", ref: "AF2" },
    "[2]": { type: "cell", ref: "AG2" },
  },
  11: {
    "[1]": { type: "cell", ref: "DG2" },
    "[2]": { type: "cell", ref: "DI2" },
  },
  12: {
    "[1]": { type: "cell", ref: "AH2" },
  },
};

const NEW_TOOLS_MAPPING: SlideMapping = {
  1: {
    "NAME OF PROJECT": { type: "cell", ref: "F2" },
    "TYPE OF PROJECT": { type: "cell", ref: "K2" },
  },
  3: {
    "[1]": { type: "cell", ref: "BZ2" },
  },
  4: {
    "[1]": { type: "cell", ref: "I2" },
    "[2]": { type: "const", value: "N/A" },
    "[3]": { type: "cell", ref: "G2" },
    "[4]": { type: "cell", ref: "H2" },
    "[5]": { type: "cell", ref: "J2" },
  },
  5: {
    "[1]": { type: "cell", ref: "CA2" },
    "[2]": { type: "const", value: "N/A" },
  },
  6: {
    "[1]": { type: "const", value: "N/A" },
    "[2]": { type: "const", value: "N/A" },
    "[3]": { type: "const", value: "N/A" },
    "[4]": { type: "const", value: "N/A" },
  },
  7: {
    "[1]": { type: "const", value: "N/A" },
  },
  8: {
    "[1]": { type: "cell", ref: "BW2" },
    "[2]": { type: "join", refs: ["CD2", "CE2"], joinWith: " " },
    "[3]": { type: "cell", ref: "CC2"
