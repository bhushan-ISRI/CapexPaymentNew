import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useEffect, useState } from "react";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Swal from "sweetalert2";
import logo from "../assets/sona-comstarlogo.png";

interface IProps {
  context: any;
  itemId: number;
  onClose?: () => void;
}
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

// Library names kept as constants so requestor docs and UTR docs can never
// accidentally be pointed at the same library again.
const REQUESTOR_DOCS_LIBRARY = "CapexPaymentDocs";
const UTR_DOCS_LIBRARY = "CapexPaymentUTRDocs";

const APperformerAdvanceFormForUTR: React.FC<IProps> = ({
  context,
  itemId,
  onClose,
}) => {
  const actionLock = React.useRef(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const sp = spfi().using(SPFx(context));
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [attachments, setAttachments] = useState<any[]>([]);
  // Files already saved in CapexPaymentUTRDocs (loaded on open, refreshed after upload)
  const [savedUtrAttachments, setSavedUtrAttachments] = useState<any[]>([]);
  // Files picked locally but not yet uploaded
  const [utrFiles, setUtrFiles] = useState<File[]>([]);
  const today = new Date();
  const localDate: string = new Date(
    today.getTime() - today.getTimezoneOffset() * 60000,
  )
    .toISOString()
    .split("T")[0];
  const [employee, setEmployee] = useState<any>({});
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");

  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);

  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [UTRRemarks, setUTRRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;
      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode",
          "EmployeeName",
          "Division",
          "Location",
          "EmployeeEmail",
          "ReportingManager/Title",
          "HOD/Title",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();
      if (user.length > 0) {
        setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };

  const getPreviousAdvances = async (vendorId: number) => {
    try {
      if (!vendorId) {
        setPreviousAdvances([]);
        return;
      }
      const data = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",
          "PaidAmount",
          "Status",
          "VendorCode/Id",
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

  // Requestor attachments — ALWAYS from CapexPaymentDocs
  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const libraryRootFolder = await sp.web.lists
        .getByTitle(REQUESTOR_DOCS_LIBRARY)
        .rootFolder();
      const path = `${libraryRootFolder.ServerRelativeUrl}/${safe}`;
      const files = await sp.web.getFolderByServerRelativePath(path).files();
      setAttachments(files || []);
    } catch (error) {
      console.log(`No requestor attachments found in ${REQUESTOR_DOCS_LIBRARY}`, error);
      setAttachments([]);
    }
  };

  // UTR attachments already saved — ALWAYS read from CapexPaymentUTRDocs ONLY.
  // Folder may not exist yet on first UTR submission, which is expected.
  const getSavedUTRAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const libraryRootFolder = await sp.web.lists
        .getByTitle(UTR_DOCS_LIBRARY)
        .rootFolder();
      const utrFolderPath = `${libraryRootFolder.ServerRelativeUrl}/${safe}`;
      const files = await sp.web.getFolderByServerRelativePath(utrFolderPath).files();
      setSavedUtrAttachments(files || []);
    } catch (error) {
      console.log(`No UTR attachments found yet in ${UTR_DOCS_LIBRARY}`, error);
      setSavedUtrAttachments([]);
    }
  };

  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName")();
      setVendors(data);
    } catch (error) {
      console.error("Vendor fetch error:", error);
    }
  };

  const getItemById = async (itemId: any) => {
    try {
      const item = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemId)
        .select(
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "Email",
          "EmployeeCode",
          "ContactNo",
          "EmployeeStatus",
          "Division",
          "Location",
          "RM",
          "HOD",
          "VendorCode/Id",
          "VendorCode/Title",
          "VendorName",
          "PODate",
          "POPaymentTerms",
          "POAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "PONumber",
          "RequestedAmountforPayment",
          "Status",
          "CurrentApproverId",
          "CapexId",
          "VoucherDate",
          "VoucherNumber",
          "ApproverRemarks",
          "InstallationDetails",
          "FinalPaymentAgainstPO",
          "ApprovalMatrix",
          "WorkflowHistory",
          "RequestorName",
        )();

      setItemData(item);

      const vendorId = item?.VendorCode?.Id || null;
      setSelectedVendorId(vendorId);
      setSelectedVendorName(item?.VendorName || "");

      if (item.CapexId) {
        await getAttachments(item.CapexId);           // -> CapexPaymentDocs
        await getSavedUTRAttachments(item.CapexId);    // -> CapexPaymentUTRDocs
      }

      if (item.ApprovalMatrix) {
        try {
          const parsed =
            typeof item.ApprovalMatrix === "string"
              ? JSON.parse(item.ApprovalMatrix)
              : item.ApprovalMatrix;
          setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          setApprovalMatrix([]);
        }
      } else {
        setApprovalMatrix([]);
      }

      if (item.WorkflowHistory) {
        try {
          const parsed =
            typeof item.WorkflowHistory === "string"
              ? JSON.parse(item.WorkflowHistory)
              : item.WorkflowHistory;
          setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          setWorkflowHistory([]);
        }
      } else {
        setWorkflowHistory([]);
      }
    } catch (error) {
      console.error("Fetch error:", error);
    }
  };

  useEffect(() => {
    if (!context || !itemId) return;
    const loadData = async () => {
      await getLoggedInUser();
      await getVendors();
      await getItemById(itemId);
    };
    void loadData();
  }, [context, itemId]);

  useEffect(() => {
    if (selectedVendorId) {
      void getPreviousAdvances(selectedVendorId);
    }
  }, [selectedVendorId]);

  // Upload UTR attachments to CapexPaymentUTRDocs ONLY, under a folder named by capexId.
  // Never touches CapexPaymentDocs.
  const uploadUTRAttachments = async (capexId: string) => {
    if (!utrFiles || utrFiles.length === 0) return;
    try {
      const safe = capexId.replace(/\//g, "_");

      const libraryRootFolder = await sp.web.lists
        .getByTitle(UTR_DOCS_LIBRARY)
        .rootFolder();

      const libraryServerRelativeUrl = libraryRootFolder.ServerRelativeUrl;
      const utrFolderPath = `${libraryServerRelativeUrl}/${safe}`;

      // Ensure the folder exists in CapexPaymentUTRDocs — create it if not
      let folderExists = true;
      try {
        await sp.web.getFolderByServerRelativePath(utrFolderPath)();
      } catch {
        folderExists = false;
      }
      if (!folderExists) {
        await sp.web
          .getFolderByServerRelativePath(libraryServerRelativeUrl)
          .folders.addUsingPath(safe);
      }

      // Upload each file into the CapexPaymentUTRDocs/<capexId> folder
      for (const file of utrFiles) {
        const arrayBuffer = await file.arrayBuffer();
        await sp.web
          .getFolderByServerRelativePath(utrFolderPath)
          .files.addUsingPath(file.name, arrayBuffer, { Overwrite: true });
      }

      // Refresh the saved list from CapexPaymentUTRDocs so UI reflects what's actually stored
      await getSavedUTRAttachments(capexId);
      // Clear the local pending-upload queue now that they're persisted
      setUtrFiles([]);
    } catch (error) {
      console.error(`UTR attachment upload error (${UTR_DOCS_LIBRARY}):`, error);
      throw error;
    }
  };

  const handleRemoveUTRFile = (index: number) => {
    const updated = [...utrFiles];
    updated.splice(index, 1);
    setUtrFiles(updated);
  };

  // Approve (Paid)
  const handleApprove = async () => {
    if (actionLock.current) return;
    if (isSubmitting) return;
    setIsSubmitting(true);
    try {
      if (!UTRDate || UTRDate.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation", text: "Please enter UTR Date.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      if (UTRDate > localDate) {
        await Swal.fire({ icon: "warning", title: "Validation", text: "UTR date cannot be a future date.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      if (!UTRNumber || UTRNumber.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation", text: "Please enter UTR Number.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      if (!UTRRemarks || UTRRemarks.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation Error", text: "Please enter UTR Remarks.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      // approverRemarks is optional here, but if something was typed it must
      // not be pure whitespace (e.g. user hits spacebar a few times).
      if (approverRemarks && approverRemarks.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation Error", text: "Approver Remarks cannot contain only spaces.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }

      // Normalize to trimmed values so nothing whitespace-padded gets persisted.
      const trimmedUTRNumber = UTRNumber.trim();
      const trimmedUTRRemarks = UTRRemarks.trim();
      const trimmedApproverRemarks = approverRemarks.trim();

      // Upload UTR attachments to CapexPaymentUTRDocs (if any new files were picked)
      if (utrFiles.length > 0) {
        try {
          await uploadUTRAttachments(itemData.CapexId);
        } catch {
          await Swal.fire({ icon: "error", title: "Upload Failed", text: `UTR attachment upload to ${UTR_DOCS_LIBRARY} failed. Please try again.`, confirmButtonText: "OK" });
          setIsSubmitting(false);
          return;
        }
      }

      const history = itemData.WorkflowHistory
        ? typeof itemData.WorkflowHistory === "string"
          ? JSON.parse(itemData.WorkflowHistory)
          : itemData.WorkflowHistory
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Paid",
        Comment: trimmedUTRRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: trimmedApproverRemarks,
          UTRDate: UTRDate ? new Date(UTRDate) : null,
          UTRNumber: trimmedUTRNumber,
          UTRRemarks: trimmedUTRRemarks,
          Status: "Paid",
          CurrentApproverId: null,
          PendingWth: "",
          WorkflowHistory: JSON.stringify(history),
        });

      await Swal.fire({ icon: "success", title: "Success", text: "Payment marked as Paid successfully.", confirmButtonText: "OK" });
      if (onClose) onClose();
    } catch (error) {
      console.error("Approve error:", error);
      await Swal.fire({ icon: "error", title: "Error", text: "An error occurred while processing the payment.", confirmButtonText: "OK" });
    } finally {
      setIsSubmitting(false);
    }
  };

  // Send Back
  const handleSendBack = async () => {
    if (actionLock.current) return;
    if (isSubmitting) return;
    setIsSubmitting(true);
    try {
      if (!UTRRemarks || UTRRemarks.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation Error", text: "Please enter UTR Remarks.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      if (approverRemarks && approverRemarks.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation Error", text: "Approver Remarks cannot contain only spaces.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      const trimmedUTRRemarksSB = UTRRemarks.trim();
      const trimmedApproverRemarksSB = approverRemarks.trim();

      const history = itemData.WorkflowHistory
        ? typeof itemData.WorkflowHistory === "string"
          ? JSON.parse(itemData.WorkflowHistory)
          : itemData.WorkflowHistory
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: trimmedUTRRemarksSB,
        Date: new Date().toISOString(),
      });

      const flow = itemData.ApprovalMatrix ? JSON.parse(itemData.ApprovalMatrix) : [];
      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);
      if (currentIndex !== -1) flow[currentIndex].Status = "Send Back";

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: trimmedApproverRemarksSB,
          Status: "Send Back",
          WorkflowHistory: JSON.stringify(history),
        });

      await Swal.fire({ icon: "success", title: "Success", text: "Request sent back successfully.", confirmButtonText: "OK" });
      if (onClose) onClose();
    } catch (error) {
      console.error(error);
      await Swal.fire({ icon: "error", title: "Error", text: "An error occurred while sending the request back.", confirmButtonText: "OK" });
    } finally {
      setIsSubmitting(false);
    }
  };

  // Reject
  const handleReject = async () => {
    if (actionLock.current) return;
    if (isSubmitting) return;
    setIsSubmitting(true);
    try {
      if (!UTRRemarks || UTRRemarks.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation Error", text: "Please enter UTR Remarks.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      if (approverRemarks && approverRemarks.trim() === "") {
        await Swal.fire({ icon: "warning", title: "Validation Error", text: "Approver Remarks cannot contain only spaces.", confirmButtonText: "OK" });
        setIsSubmitting(false);
        return;
      }
      const trimmedUTRRemarksRej = UTRRemarks.trim();
      const trimmedApproverRemarksRej = approverRemarks.trim();

      const history = itemData.WorkflowHistory
        ? typeof itemData.WorkflowHistory === "string"
          ? JSON.parse(itemData.WorkflowHistory)
          : itemData.WorkflowHistory
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Rejected",
        Comment: trimmedUTRRemarksRej,
        Date: new Date().toISOString(),
      });

      const flow = itemData.ApprovalMatrix ? JSON.parse(itemData.ApprovalMatrix) : [];
      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);
      if (currentIndex !== -1) flow[currentIndex].Status = "Reject";

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: trimmedApproverRemarksRej,
          Status: "Reject",
          CurrentApproverId: null,
          PendingWth: "",
          WorkflowHistory: JSON.stringify(history),
        });

      await Swal.fire({ icon: "success", title: "Success", text: "Request rejected successfully.", confirmButtonText: "OK" });
      if (onClose) onClose();
    } catch (error) {
      console.error(error);
      await Swal.fire({ icon: "error", title: "Error", text: "An error occurred while rejecting the request.", confirmButtonText: "OK" });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleExit = () => {
    if (onClose) onClose();
  };

  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1> Advance Payment (Approver) </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  <li className="approval-step">
                    {`Initiator`} - {itemData?.EmployeeName}
                  </li>
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${
                        a.Status === "In Progress"
                          ? "active"
                          : a.Status === "Approved"
                            ? "approved"
                            : a.Status === "Reject"
                              ? "reject"
                              : a.Status === "Send Back"
                                ? "sendback"
                                : ""
                      }`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div className="borderedbox">
              {/* Requestor Information */}
              <div className="heading1">
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">Employee Code</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">Employee Name</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">Employee Email</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.Email}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">Contact No</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">Employee Status</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.EmployeeStatus}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">Division</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">Location</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">RM</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.RM}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">HOD</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.HOD}</label>
                  </div>
                </div>
              </div>

              {/* Vendor & PO Details */}
              <div className="heading1">
                <label>Vendor & PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.VendorCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.VendorName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.PONumber}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData.PODate ? new Date(itemData.PODate).toLocaleDateString("en-GB") : ""}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Terms</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.POPaymentTerms}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.POAmount}</label>
                  </div>
                </div>
              </div>

              {/* Approver Action */}
              <div className="heading1">
                <label>Approver Action</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                    <input value={itemData.ApproverRemarks || ""} className="font-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input
                      value={itemData.VoucherDate ? new Date(itemData.VoucherDate).toLocaleDateString("en-GB") : ""}
                      className="font-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input value={itemData.VoucherNumber || ""} className="font-control readonly" readOnly />
                  </div>
                </div>
              </div>

              {/* MRN & Payment Details */}
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>MRN & Payment Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Number</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData?.MRNNumber}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNDtae ? new Date(itemData.MRNDtae).toLocaleDateString("en-GB") : ""}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData?.MRNAmountwithGST}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData?.RequestedAmountforPayment}</label>
                  </div>
                </div>
              </div>

              {/* Upload Document - Requestor Attachments, from CapexPaymentDocs */}
              <div className="heading1">
                <label>Upload Document</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length === 0 ? (
                      <p>No attachments</p>
                    ) : (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                              {file.Name}
                            </a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </div>

              {/* UTR Details Section — everything here writes to / reads from CapexPaymentUTRDocs only */}
              <div className="heading1">
                <label>UTR Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">UTR Date</label>
                    <input
                      type="date"
                      className="font-control"
                      value={UTRDate}
                      max={new Date().toISOString().split("T")[0]}
                      onChange={(e) => setUTRDate(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">UTR Number</label>
                    <input
                      value={UTRNumber}
                      className="font-control"
                      onChange={(e) => setUTRNumber(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">UTR Remarks</label>
                    <input
                      className="font-control"
                      value={UTRRemarks}
                      onChange={(e) => setUTRRemarks(e.target.value)}
                    />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-8">
                    <label className="font">UTR Attachments</label>
                    <input
                      type="file"
                      multiple
                      className="font-control"
                      onChange={(e) => {
                        if (e.target.files) {
                          setUtrFiles((prev) => [
                            ...prev,
                            ...Array.from(e.target.files!),
                          ]);
                        }
                      }}
                    />

                    {/* Files already persisted in CapexPaymentUTRDocs */}
                    {savedUtrAttachments.length > 0 && (
                      <>
                        <p style={{ margin: "8px 0 2px", fontSize: "12px", fontWeight: 600 }}>
                          Already saved:
                        </p>
                        <ul style={{ marginBottom: "8px" }}>
                          {savedUtrAttachments.map((file: any, index: number) => (
                            <li key={`saved-${index}`}>
                              <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                                {file.Name}
                              </a>
                            </li>
                          ))}
                        </ul>
                      </>
                    )}

                    {/* Files picked but not yet uploaded */}
                    {utrFiles.length > 0 && (
                      <>
                        <p style={{ margin: "8px 0 2px", fontSize: "12px", fontWeight: 600 }}>
                          Pending upload:
                        </p>
                        <ul style={{ marginTop: "0", paddingLeft: "0", listStyle: "none" }}>
                          {utrFiles.map((file: File, index: number) => (
                            <li
                              key={index}
                              style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "4px" }}
                            >
                              <a
                                href={URL.createObjectURL(file)}
                                target="_blank"
                                rel="noopener noreferrer"
                                style={{ fontSize: "13px" }}
                              >
                                {file.name}
                              </a>
                              <button
                                type="button"
                                style={{
                                  color: "red",
                                  background: "none",
                                  border: "none",
                                  cursor: "pointer", 
                                  fontSize: "12px",
                                  padding: "0",
                                }}
                                onClick={() => handleRemoveUTRFile(index)}
                              >
                                Remove
                              </button>
                            </li>
                          ))}
                        </ul>
                      </>
                    )}
                  </div>
                </div>
              </div>

              {/* Workflow History */}
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <div className="workflow-history">
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
                              .filter(
                                (h: any) =>
                                  h.ActionTaken &&
                                  h.ActionTaken !== "Draft Saved" &&
                                  h.ActionTaken !== "Edited",
                              )
                              .map((h: any, idx: number) => (
                                <tr key={idx}>
                                  <td style={{ padding: "8px" }}>{h.CurrentApprover || ""}</td>
                                  <td style={{ padding: "8px" }}>{h.ActionTaken || ""}</td>
                                  <td style={{ padding: "8px" }}>
                                    {h.Date ? new Date(h.Date).toLocaleDateString("en-GB") : ""}
                                  </td>
                                  <td style={{ padding: "8px" }}>{h.Comment || ""}</td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Action Buttons */}
              <div className="row my-3">
                <div className="col-md-12">
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "5px" }}>
                    <a
                      className={`submit-btn ${isSubmitting ? "disabled-btn" : ""}`}
                      onClick={!isSubmitting ? handleApprove : undefined}
                    >
                      {isSubmitting ? "Processing..." : "Paid"}
                    </a>
                    <a
                      className={`Rework-btn ${isSubmitting ? "disabled-btn" : ""}`}
                      onClick={!isSubmitting ? handleSendBack : undefined}
                    >
                      {isSubmitting ? "Processing..." : "Send Back"}
                    </a>
                    <a
                      className={`Reject-btn ${isSubmitting ? "disabled-btn" : ""}`}
                      onClick={!isSubmitting ? handleReject : undefined}
                    >
                      {isSubmitting ? "Processing..." : "Reject"}
                    </a>
                    <a href="#" onClick={handleExit} className="reset-btn">
                      Exit
                    </a>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default APperformerAdvanceFormForUTR;