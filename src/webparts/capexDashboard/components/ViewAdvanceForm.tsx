import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import 'bootstrap/dist/css/bootstrap.min.css';
import logo from "../assets/sona-comstarlogo.png";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

interface IPreviousAdvance {
  PONumber: string;
  RequestAdvanceAmount: string;
  Created: string;
  VoucherDate: string;
  VouchingNumber: string;
  PaidAmount: string;
  Status: string;
}

// Library names kept as constants so requestor docs and UTR docs can never
// accidentally be pointed at the same library again.
const REQUESTOR_DOCS_LIBRARY = "CapexPaymentDocs";
const UTR_DOCS_LIBRARY = "CapexPaymentUTRDocs";

const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  // Requestor attachments (root folder files only) - from CapexPaymentDocs
  const [attachments, setAttachments] = useState<any[]>([]);
  // UTR attachments - from CapexPaymentUTRDocs ONLY
  const [utrAttachments, setUtrAttachments] = useState<any[]>([]);

  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = useState<any>({});
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<IPreviousAdvance[]>([]);
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [requesterRemarks, setRequesterRemarks] = useState("");
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VouchingNumber, setVouchingNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [UTRRemarks, setUTRRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const getPreviousAdvances = async (vendorId: number) => {
    try {
      if (!vendorId) { setPreviousAdvances([]); return; }
      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber", "RequestAdvanceAmount", "Created",
          "VoucherDate", "VouchingNumber", "PaidAmount", "Status", "VendorCode/Id",
        )
        .expand("VendorCode")
        .filter(`VendorCode/Id eq ${vendorId} and Status eq 'Paid'`)
        .orderBy("Created", false)();
      setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      setPreviousAdvances([]);
    }
  };

  // Requestor attachments — ALWAYS from CapexPaymentDocs, never from the UTR library
  const getAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const libraryRootFolder = await sp.web.lists
        .getByTitle(REQUESTOR_DOCS_LIBRARY)
        .rootFolder();
      const folderPath = `${libraryRootFolder.ServerRelativeUrl}/${safeCapexId}`;
      const files = await sp.web.getFolderByServerRelativePath(folderPath).files();
      setAttachments(files || []);
    } catch (error) {
      console.log(`No requestor attachments found in ${REQUESTOR_DOCS_LIBRARY} for ${capexId}`, error);
      setAttachments([]);
    }
  };

  // UTR attachments — ALWAYS from CapexPaymentUTRDocs, never from the requestor docs library.
  // This folder only exists once the AP Performer has uploaded UTR proof, so a missing
  // folder is expected/normal and is swallowed silently (empty list shown in UI).
  const getUTRAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;
      const safeCapexId = capexId.replace(/\//g, "_");

      const libraryRootFolder = await sp.web.lists
        .getByTitle(UTR_DOCS_LIBRARY)
        .rootFolder();

      const utrFolderPath = `${libraryRootFolder.ServerRelativeUrl}/${safeCapexId}`;
      const files = await sp.web.getFolderByServerRelativePath(utrFolderPath).files();
      setUtrAttachments(files || []);
    } catch (error) {
      // Folder won't exist until AP Performer uploads UTR docs — safe to ignore
      console.log(`No UTR attachments found in ${UTR_DOCS_LIBRARY} for ${capexId}`, error);
      setUtrAttachments([]);
    }
  };

  const getVendors = async () => {
    try {
      const data = await sp.web.lists.getByTitle("VendorMaster").items.select("Id", "VendorCode", "VendorName")();
      setVendors(data);
    } catch (error) { console.error("Vendor fetch error:", error); }
  };

  const getEmployeeDetails = async () => {
    try {
      if (!formData?.Email) return;
      const user = await sp.web.lists.getByTitle("EmployeeMaster").items
        .select("EmployeeCode", "EmployeeName", "Division", "Location", "EmployeeEmail",
          "ReportingManager/Title", "HOD/Title", "ContactNo", "EmployeeStatus", "CostCenter")
        .expand("ReportingManager", "HOD").filter(`EmployeeEmail eq '${formData.Email}'`).top(1)();
      if (user.length > 0) setEmployee(user[0]);
    } catch (error) { console.log("Error fetching employee:", error); }
  };

  useEffect(() => {
    if (!formData) return;
    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POPaymentTerms || "");
    setPoAmount(formData.POAmount || "");
    setSelectedVendorCode(formData.VendorCode || "");
    setSelectedVendorName(formData.VendorName || "");
    setMrnNumber(formData.MRNNumber || "");
    setMrnDate(formData.MRNDtae?.split("T")[0] || "");
    setMrnAmount(formData.MRNAmountwithGST || "");
    setRequestedAmount(formData.RequestedAmountforPayment || "");
    setFinalPayment(formData.FinalPaymentAgainstPO ? "Yes" : "No");
    setInstallationDetails(formData.InstallationDetails || "");
    setRequesterRemarks(formData.RequesterRemarks || "");
    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVouchingNumber(formData.VoucherNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");
    setUTRRemarks(formData.UTRRemarks || "");

    if (formData.CapexId) {
      void getAttachments(formData.CapexId);     // -> CapexPaymentDocs
      void getUTRAttachments(formData.CapexId);  // -> CapexPaymentUTRDocs
    }

    if (formData?.ApprovalMatrix) {
      try {
        const parsed = typeof formData.ApprovalMatrix === "string" ? JSON.parse(formData.ApprovalMatrix) : formData.ApprovalMatrix;
        setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
      } catch { setApprovalMatrix([]); }
    } else { setApprovalMatrix([]); }

    if (formData?.WorkflowHistory) {
      try {
        const parsed = typeof formData.WorkflowHistory === "string" ? JSON.parse(formData.WorkflowHistory) : formData.WorkflowHistory;
        setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
      } catch { setWorkflowHistory([]); }
    } else { setWorkflowHistory([]); }
  }, [formData]);

  useEffect(() => {
    if (vendors.length > 0 && selectedVendorCode) {
      const match = vendors.find((v) => v.VendorCode === selectedVendorCode);
      if (match) {
        setSelectedVendorId(match.Id);
        void getPreviousAdvances(match.Id);
      }
    }
  }, [vendors, selectedVendorCode]);

  useEffect(() => { void getEmployeeDetails(); void getVendors(); }, []);

  const handleExit = () => { if (onClose) onClose(); else window.location.reload(); };

  // ===== Ribbon color logic =====
  // Colors: "approved" = green, "active" = orange (current approver), "upcoming" = yellow,
  // "rejected" = red. Driven by the overall request Status (formData.Status), the
  // per-step Status inside approvalMatrix, and — for Send Back, where CurrentApprover
  // is cleared — the WorkflowHistory entries in the CapexPayment list.
  //
  // Rules:
  // - Paid: Initiator + every approver step = green.
  // - Reject: the step that has Status "Reject" = red; steps before it = green;
  //   steps after it = yellow.
  // - Send Back: CurrentApprover goes blank, so we can't trust matrix "In Progress".
  //   The Initiator (requester) = orange, every approver step = yellow.
  // - Otherwise (still in progress, e.g. Pending for Approval / Pending for Vouching
  //   Update / Pending for UTR Update): steps already Approved = green, the step
  //   with Status "In Progress" = orange, steps after that = yellow.
  const overallStatus: string = formData?.Status || "";

  const buildRibbonSteps = () => {
    const initiatorStep = {
      Role: "Initiator",
      Name: formData?.EmployeeName || employee.EmployeeName || "",
      Status: "Approved",
    };
    const approverSteps = approvalMatrix.filter((a) => a.Role !== "Initiator");
    const steps = [initiatorStep, ...approverSteps];

    if (overallStatus === "Paid") {
      return steps.map((s) => ({ ...s, _color: "approved" }));
    }

    if (overallStatus === "Reject") {
      const rejectIndex = steps.findIndex((s) => s.Status === "Reject" || s.Status === "Rejected");
      return steps.map((s, idx) => {
        if (rejectIndex === -1) return { ...s, _color: "" };
        if (idx === rejectIndex) return { ...s, _color: "rejected" };
        if (idx < rejectIndex) return { ...s, _color: "approved" };
        return { ...s, _color: "upcoming" };
      });
    }

    if (overallStatus === "Send Back" || overallStatus === "Draft") {
      // CurrentApprover is blank by this point for Send Back, and for Draft the
      // request hasn't been submitted yet — either way the requester is the
      // active step (orange), not "done" (green), and every approver step is
      // upcoming (yellow).
      return steps.map((s) =>
        s.Role === "Initiator" ? { ...s, _color: "active" } : { ...s, _color: "upcoming" },
      );
    }

    // Default: still in progress through the approval chain.
    return steps.map((s) => {
      if (s.Status === "Approved") return { ...s, _color: "approved" };
      if (s.Status === "In Progress") return { ...s, _color: "active" };
      return { ...s, _color: "upcoming" };
    });
  };

  const getStepClass = (color: string) => {
    switch (color) {
      case "approved": return "approved";   // green
      case "active": return "active";       // orange — current approver / send-back target
      case "upcoming": return "upcoming";   // yellow — not yet reached
      case "rejected": return "rejected";   // red
      default: return "";
    }
  };
  // Add these rules to advanced.scss if not already present:
  //   .approval-step.approved { background-color: green; color: #fff; }
  //   .approval-step.active   { background-color: orange; color: #fff; }
  //   .approval-step.upcoming { background-color: #f5d800; color: #333; }
  //   .approval-step.rejected { background-color: red; color: #fff; }

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row"><div className="col-md-12"><div className="Main-Boxpoup">
        <div className="bordered"><img src={logo} /><h1>Advance Payment (View)</h1></div>

        {approvalMatrix.length === 0 ? (<p>No approval data</p>) : (
          <div className="displayWF">
            <ul className="approval-flow">
              {buildRibbonSteps().map((a, index) => (
                <li key={index} className={`approval-step ${getStepClass(a._color)}`}>
                  {a.Role} - {a.Name}
                </li>
              ))}
            </ul>
          </div>
        )}

        <div className="borderedbox">
          <div className="heading1"><label>Requestor Information</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Employee Code</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeCode}</label></div>
              <div className="col-md-4"><label className="font">Employee Name</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeName}</label></div>
              <div className="col-md-4"><label className="font">Employee Email</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeEmail}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Contact No</label> : &nbsp;&nbsp;<label className="fonttext">{employee.ContactNo}</label></div>
              <div className="col-md-4"><label className="font">Employee Status</label> : &nbsp;&nbsp;<label className="fonttext">{employee.EmployeeStatus}</label></div>
              <div className="col-md-4"><label className="font">Division</label> : &nbsp;&nbsp;<label className="fonttext">{employee.Division}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Location</label> : &nbsp;&nbsp;<label className="fonttext">{employee.Location}</label></div>
              <div className="col-md-4"><label className="font">RM</label> : &nbsp;&nbsp;<label className="fonttext">{employee.ReportingManager?.Title}</label></div>
              <div className="col-md-4"><label className="font">HOD</label> : &nbsp;&nbsp;<label className="fonttext">{employee.HOD?.Title}</label></div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Vendor & PO</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Vendor Code</label> : &nbsp;&nbsp;<label className="fonttext">{selectedVendorCode}</label></div>
              <div className="col-md-4"><label className="font">Vendor Name</label> : &nbsp;&nbsp;<label className="fonttext">{selectedVendorName}</label></div>
              <div className="col-md-4"><label className="font">PO Number</label> : &nbsp;&nbsp;<label className="fonttext">{poNumber}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">PO Date</label> : &nbsp;&nbsp;<label className="fonttext">{poDate}</label></div>
              <div className="col-md-4"><label className="font">PO Amount</label> : &nbsp;&nbsp;<label className="fonttext">{poAmount}</label></div>
              <div className="col-md-4"><label className="font">PO Payment Terms</label> : &nbsp;&nbsp;<label className="fonttext">{poTerms}</label></div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>MRN & Payment Details</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">MRN Number</label> : &nbsp;&nbsp;<label className="fonttext">{mrnNumber}</label></div>
              <div className="col-md-4"><label className="font">MRN Date</label> : &nbsp;&nbsp;<label className="fonttext">{mrnDate}</label></div>
              <div className="col-md-4"><label className="font">MRN Amount</label> : &nbsp;&nbsp;<label className="fonttext">{mrnAmount}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Requested Amount</label> : &nbsp;&nbsp;<label className="fonttext">{requestedAmount}</label></div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Previous Advances</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-12">
                <div style={{ overflowX: "auto" }}>
                  <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                    <thead className="text-white" style={{ backgroundColor: "rgb(60, 62, 69)" }}>
                      <tr>
                        <th className="px-4 py-2">PO Number</th>
                        <th className="px-4 py-2">Previous Advance</th>
                        <th className="px-4 py-2">Requested Date</th>
                        <th className="px-4 py-2">Paid Date</th>
                        <th className="px-4 py-2">Voucher No</th>
                        <th className="px-4 py-2">Settled Amount</th>
                        <th className="px-4 py-2">Pending Advance</th>
                      </tr>
                    </thead>
                    <tbody>
                      {previousAdvances.length === 0 ? (
                        <tr>
                          <td colSpan={7} style={{ textAlign: "center", padding: "10px" }}>No previous advances available</td>
                        </tr>
                      ) : (
                        previousAdvances.map((item: any, index: number) => {
                          const pending = Math.max(0, Number(item.RequestAdvanceAmount || 0) - Number(item.PaidAmount || 0));
                          return (
                            <tr key={index}>
                              <td className="px-4 py-2">{item.PONumber}</td>
                              <td className="px-4 py-2">{item.RequestAdvanceAmount}</td>
                              <td className="px-4 py-2">{item.Created ? new Date(item.Created).toLocaleDateString("en-GB") : ""}</td>
                              <td className="px-4 py-2">{item.VoucherDate ? new Date(item.VoucherDate).toLocaleDateString("en-GB") : ""}</td>
                              <td className="px-4 py-2">{item.VouchingNumber}</td>
                              <td className="px-4 py-2">{item.PaidAmount}</td>
                              <td className="px-4 py-2">{pending}</td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Final Payment Details</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">Final Payment Against PO</label> : &nbsp;&nbsp;<label className="fonttext">{finalPayment}</label></div>
            </div>
            {finalPayment === "Yes" && (
              <div className="row mb-20">
                <div className="col-md-6"><label className="font">Installation Details</label> : &nbsp;&nbsp;<label className="fonttext">{installationDetails}</label></div>
              </div>
            )}
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Requester Remarks</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-6"><label className="font">Requester Remarks</label> : &nbsp;&nbsp;<label className="fonttext">{requesterRemarks}</label></div>
            </div>
          </div>

          {/* Requestor Attachments — root folder only, from CapexPaymentDocs */}
          <div className="heading1" style={{ marginTop: "10px" }}><label>Upload Document</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4">
                <label className="font">Attachments</label>
                {attachments.length > 0 ? (
                  <ul>{attachments.map((file: any, index: number) => (
                    <li key={index}><a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">{file.Name}</a></li>
                  ))}</ul>
                ) : <p>No attachments</p>}
              </div>
            </div>
          </div>

          {/* Voucher Details */}
          <div className="main-formcontainer" style={{ marginTop: "10px" }}>
            <div className="row mb-20">
              <div className="col-md-6"><label className="font">Voucher Date</label> : &nbsp;&nbsp;<label className="fonttext">{voucherDate}</label></div>
              <div className="col-md-6"><label className="font">Voucher Number</label> : &nbsp;&nbsp;<label className="fonttext">{VouchingNumber}</label></div>
            </div>
          </div>

          {/* UTR Details — UTR fields + UTR Attachments shown together, sourced ONLY from CapexPaymentUTRDocs */}
          <div className="heading1" style={{ marginTop: "10px" }}><label>UTR Details</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-4"><label className="font">UTR Date</label> : &nbsp;&nbsp;<label className="fonttext">{UTRDate}</label></div>
              <div className="col-md-4"><label className="font">UTR Number</label> : &nbsp;&nbsp;<label className="fonttext">{UTRNumber}</label></div>
              <div className="col-md-4"><label className="font">UTR Remarks</label> : &nbsp;&nbsp;<label className="fonttext">{UTRRemarks}</label></div>
            </div>
            <div className="row mb-20">
              <div className="col-md-4">
                <label className="font">UTR Attachments</label>
                {utrAttachments.length > 0 ? (
                  <ul>{utrAttachments.map((file: any, index: number) => (
                    <li key={index}>
                      <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                        {file.Name}
                      </a>
                    </li>
                  ))}</ul>
                ) : <p>No UTR attachments</p>}
              </div>
            </div>
          </div>

          <div className="heading1" style={{ marginTop: "10px" }}><label>Workflow History</label></div>
          <div className="main-formcontainer">
            <div className="row mb-20">
              <div className="col-md-12">
                {workflowHistory.length === 0 ? (<p>No history available</p>) : (
                  <table className="workflow-table" style={{ width: "100%" }}>
                    <thead>
                      <tr>
                        <th style={{ padding: "8px", textAlign: "left" }}>Action By</th>
                        <th style={{ padding: "8px", textAlign: "left" }}>Action Taken</th>
                        <th style={{ padding: "8px", textAlign: "left" }}>Date</th>
                        <th style={{ padding: "8px", textAlign: "left" }}>Comment</th>
                      </tr>
                    </thead>
                    <tbody>
                      {workflowHistory
                        .filter((h: any) => h.ActionTaken && h.ActionTaken !== "Edited")
                        .map((h: any, idx: number) => (
                          <tr key={idx}>
                            <td style={{ padding: "8px" }}>{h.CurrentApprover || ""}</td>
                            <td style={{ padding: "8px" }}>{h.ActionTaken || ""}</td>
                            <td style={{ padding: "8px" }}>{h.Date ? new Date(h.Date).toLocaleDateString("en-GB") : ""}</td>
                            <td style={{ padding: "8px" }}>{h.Comment || ""}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          </div>

          <div className="row my-3">
            <div className="col-md-12">
              <div className="text-center">
                <a href="#" onClick={handleExit} className="reset-btn">Exit</a>
              </div>
            </div>
          </div>
        </div>
      </div></div></div>
    </div>
  );
};

export default ViewAdvanceForm;