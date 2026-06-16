import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "bootstrap/dist/css/bootstrap.min.css";
import logo from "../assets/sona-comstarlogo.png";
import Swal from "sweetalert2";
interface IProps {
  context: any;
  itemId: number;
  onClose: () => void;
}

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const APperformerAdvanceform: React.FC<IProps> = ({
  context,
  itemId,
  onClose,
}) => {
  const sp = spfi().using(SPFx(context));
  const [isSubmitting, setIsSubmitting] = useState(false);
  const actionLock = React.useRef(false);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const today = new Date();
  const localDate: string = new Date(
    today.getTime() - today.getTimezoneOffset() * 60000,
  )
    .toISOString()
    .split("T")[0];
  const [employee, setEmployee] = useState<any>({});
  const [attachments, setAttachments] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [voucherNumber, setVoucherNumber] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
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
      if (user.length > 0) setEmployee(user[0]);
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };

  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/CapexPaymentDocs/${safe}`;
      const files = await sp.web.getFolderByServerRelativePath(path).files();
      setAttachments(files);
    } catch {
      setAttachments([]);
    }
  };

  const getItemById = async (id: any) => {
    try {
      const item = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(id)
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
          "InstallationDetails",
          "FinalPaymentAgainstPO",
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
          "ApprovalMatrix",
          "WorkflowHistory",
          "RequestorName",
        )();
      setItemData(item);
      const vendorId = item?.VendorCode?.Id || null;
      setSelectedVendorId(vendorId);
      setSelectedVendorName(item?.VendorName || "");
      if (item.CapexId) await getAttachments(item.CapexId);
      if (item.ApprovalMatrix) {
        try {
          const parsed =
            typeof item.ApprovalMatrix === "string"
              ? JSON.parse(item.ApprovalMatrix)
              : item.ApprovalMatrix;
          setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
        } catch {
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
        } catch {
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
    if (selectedVendorId) void getPreviousAdvances(selectedVendorId);
  }, [selectedVendorId]);

  const handleApprove = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    if (isSubmitting) return;
    try {
      setIsSubmitting(true);
      if (!voucherDate || voucherDate.trim() === "") {
        await Swal.fire({
          icon: "warning",
          title: "Validation Error",
          text: "Please enter Voucher Date.",
          confirmButtonText: "OK",
        });
        return;
      }
      if (voucherDate > localDate) {
        await Swal.fire({
          icon: "warning",
          title: "Invalid Date",
          text: "Voucher Date cannot be a future date.",
          confirmButtonText: "OK",
        });
        return;
      }
      if (!voucherNumber || voucherNumber.trim() === "") {
        await Swal.fire({
          icon: "warning",
          title: "Validation Error",
          text: "Please enter Voucher Number.",
          confirmButtonText: "OK",
        });
        return;
      }
      if (!approverRemarks || approverRemarks.trim() === "") {
        await Swal.fire({
          icon: "warning",
          title: "Validation Error",
          text: "Please enter Remarks.",
          confirmButtonText: "OK",
        });
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];
      const currentUser = context.pageContext.user.displayName;
      const currentIndex = flow.findIndex((a: any) => a.Name === currentUser);
      if (currentIndex === -1) {
        await Swal.fire({
          icon: "error",
          title: "Access Denied",
          text: "You are not the current approver.",
          confirmButtonText: "OK",
        });
        return;
      }
      flow[currentIndex].Status = "Approved";
      let nextApproverId = null;
      if (flow[currentIndex + 1]) {
        flow[currentIndex + 1].Status = "In Progress";
        nextApproverId = flow[currentIndex + 1].Id;
      }

      const history = itemData.WorkflowHistory
        ? JSON.parse(itemData.WorkflowHistory)
        : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Approved",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          VoucherDate: voucherDate ? new Date(voucherDate) : null,
          VoucherNumber: voucherNumber,
          Status: "Pending for UTR Update",
          WorkflowHistory: JSON.stringify(history),
        });
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Vouching details submitted successfully.",
        confirmButtonText: "OK",
      });

      onClose();
    } catch (error) {
      console.error("Approve error:", error);
      await Swal.fire({
        icon: "error",
        title: "Error",
        text: "An error occurred while submitting the voucher details.",
        confirmButtonText: "OK",
      });
    } finally {
      actionLock.current = false;
      setIsSubmitting(false);
    }
  };

  const handleSendBack = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    setIsSubmitting(true);
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        await Swal.fire({
          icon: "warning",
          title: "Validation Error",
          text: "Please enter Remarks.",
          confirmButtonText: "OK",
        });
        return;
      }
      let flow: any[] = [];
      try {
        flow =
          typeof itemData.ApprovalMatrix === "string"
            ? JSON.parse(itemData.ApprovalMatrix)
            : itemData.ApprovalMatrix || [];
      } catch {
        flow = [];
      }
      if (!Array.isArray(flow) || flow.length === 0) {
        await Swal.fire({
          icon: "error",
          title: "Workflow Error",
          text: "Approval Matrix is empty.",
          confirmButtonText: "OK",
        });
        return;
      }
      const currentIndex = flow.findIndex(
        (x: any) => x.Status === "Pending" || x.Status === "In Progress",
      );
      if (currentIndex === -1) {
        await Swal.fire({
          icon: "warning",
          title: "Approver Not Found",
          text: "No current approver found.",
          confirmButtonText: "OK",
        });
        return;
      }

      flow[currentIndex].Status = "Send Back";
      let previousApproverId: number | null = null;
      if (currentIndex > 0) {
        flow[currentIndex - 1].Status = "Pending";
        previousApproverId = flow[currentIndex - 1].Id;
      }

      let history: any[] = [];
      try {
        history =
          typeof itemData.WorkflowHistory === "string"
            ? JSON.parse(itemData.WorkflowHistory)
            : itemData.WorkflowHistory || [];
      } catch {
        history = [];
      }
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
        CurrentStatus: "Send Back",
      });

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Send Back",
          WorkflowHistory: JSON.stringify(history),
        });
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Request sent back successfully.",
        confirmButtonText: "OK",
      });

      onClose();
    } catch (error) {
      console.error(error);
      await Swal.fire({
        icon: "error",
        title: "Error",
        text: "An error occurred while processing your request.",
        confirmButtonText: "OK",
      });
    } finally {
      actionLock.current = false;
      setIsSubmitting(false);
    }
  };

  const handleReject = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    setIsSubmitting(true);
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        await Swal.fire({
          icon: "warning",
          title: "Validation Error",
          text: "Please enter Remarks.",
          confirmButtonText: "OK",
        });
        return;
      }
      let flow: any[] = [];
      try {
        flow =
          typeof itemData.ApprovalMatrix === "string"
            ? JSON.parse(itemData.ApprovalMatrix)
            : itemData.ApprovalMatrix || [];
      } catch {
        flow = [];
      }
      if (!Array.isArray(flow) || flow.length === 0) {
        await Swal.fire({
          icon: "error",
          title: "Workflow Error",
          text: "Approval Matrix is empty.",
          confirmButtonText: "OK",
        });
        return;
      }

      const currentIndex = flow.findIndex(
        (x: any) => x.Status === "Pending" || x.Status === "In Progress",
      );
      if (currentIndex === -1) {
        await Swal.fire({
          icon: "warning",
          title: "Approver Not Found",
          text: "No current approver found.",
          confirmButtonText: "OK",
        });
        return;
      }

      flow[currentIndex].Status = "Reject";
      let history: any[] = [];
      try {
        history =
          typeof itemData.WorkflowHistory === "string"
            ? JSON.parse(itemData.WorkflowHistory)
            : itemData.WorkflowHistory || [];
      } catch {
        history = [];
      }
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Rejected",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
        CurrentStatus: "Reject",
      });

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Reject",
          WorkflowHistory: JSON.stringify(history),
        });
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Request rejected successfully.",
        confirmButtonText: "OK",
      });
      onClose();
    } catch (error) {
      console.error(error);
      await Swal.fire({
        icon: "error",
        title: "Error",
        text: "An error occurred while processing your request.",
        confirmButtonText: "OK",
      });
    } finally {
      actionLock.current = false;
      setIsSubmitting(false);
    }
  };

  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1>Advance Payment (Approver)</h1>
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
                      className={`approval-step ${a.Status === "In Progress" ? "active" : a.Status === "Approved" ? "approved" : a.Status === "Reject" ? "reject" : a.Status === "Send Back" ? "sendback" : ""}`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <div className="borderedbox">
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Employee Code</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Name</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Email</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{itemData.Email}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Contact No</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Status</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Division</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Location</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">RM</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.RM}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">HOD</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
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
                      {itemData.PODate
                        ? new Date(itemData.PODate).toLocaleDateString("en-GB")
                        : ""}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Terms</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData.POPaymentTerms}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.POAmount}</label>
                  </div>
                </div>
              </div>
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
                      {itemData?.MRNDtae
                        ? new Date(itemData.MRNDtae).toLocaleDateString("en-GB")
                        : ""}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNAmountwithGST}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.RequestedAmountforPayment}
                    </label>
                  </div>
                </div>
              </div>
              <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input
                      type="date"
                      value={voucherDate}
                      onChange={(e) => setVoucherDate(e.target.value)}
                      max={new Date().toISOString().split("T")[0]}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input
                      value={voucherNumber}
                      onChange={(e) => setVoucherNumber(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                    <textarea
                      value={approverRemarks}
                      className="form-control"
                      onChange={(e) => setApproverRemarks(e.target.value)}
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
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
                            <a
                              href={file.ServerRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.Name}
                            </a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <table
                        className="workflow-table"
                        style={{ width: "100%" }}
                      >
                        <thead>
                          <tr>
                            <th style={{ padding: "8px", textAlign: "left" }}>
                              Action By
                            </th>
                            <th style={{ padding: "8px", textAlign: "left" }}>
                              Action Taken
                            </th>
                            <th style={{ padding: "8px", textAlign: "left" }}>
                              Date
                            </th>
                            <th style={{ padding: "8px", textAlign: "left" }}>
                              Comment
                            </th>
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
                                <td style={{ padding: "8px" }}>
                                  {h.CurrentApprover || ""}
                                </td>
                                <td style={{ padding: "8px" }}>
                                  {h.ActionTaken || ""}
                                </td>
                                <td style={{ padding: "8px" }}>
                                  {h.Date
                                    ? new Date(h.Date).toLocaleDateString(
                                        "en-GB",
                                      )
                                    : ""}
                                </td>
                                <td style={{ padding: "8px" }}>
                                  {h.Comment || ""}
                                </td>
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
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "5px",
                    }}
                  >
                    <a
                      className={`submit-btn ${isSubmitting ? "disabled-btn" : ""}`}
                      onClick={!isSubmitting ? handleApprove : undefined}
                    >
                      {isSubmitting ? "Processing..." : "Submit"}
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
                    <a href="#" onClick={onClose} className="reset-btn">
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

export default APperformerAdvanceform;