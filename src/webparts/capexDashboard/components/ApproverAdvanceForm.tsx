import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
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

const ApproverAdvanceForm: React.FC<IProps> = ({
  context,
  itemId,
  onClose,
}) => {
  const sp = spfi().using(SPFx(context));
  const actionLock = React.useRef(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [employee, setEmployee] = React.useState<any>({});
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [approverRemarks, setApproverRemarks] = useState("");
  const [attachments, setAttachments] = useState<any[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");
  // Shown when "Whether this is the Final Payment against the PO" = No
  const [installationRequestNumber, setInstallationRequestNumber] =
    useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  const getPreviousAdvances = async (vendorId: number) => {
    try {
      if (!vendorId) {
        setPreviousAdvances([]);
        return;
      }
      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",
          "VouchingNumber",
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

  const getAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const folderPath = `CapexPaymentDocs/${safeCapexId}`;
      const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();
      setAttachments(files || []);
    } catch (error) {
      setAttachments([]);
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
          "VendorCode",
          "VendorName",
          "PODate",
          "POPaymentTerms",
          "POBasicAmount",
          "POGSTAmount",
          "POOtherAmount",
          "POAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNBasicAmount",
          "MRNGSTAmount",
          "MRNOtherAmount",
          "MRNAmountwithGST",
          "PONumber",
          "RequestedAmountforPayment",
          "Status",
          "CurrentApproverId",
          "CapexId",
          "InstallationDetails",
          "InstallationRequestNumber",
          "FinalPaymentAgainstPO",
          "ApprovalMatrix",
          "WorkflowHistory",
          "RequestorName",
        )();

      setItemData(item);
      // VendorCode on CapexPayment is a plain text field (the vendor code string),
      // NOT a SharePoint Lookup — so there's no item.VendorCode.Id to read here.
      // We resolve the real VendorMaster list-item Id by matching this text value
      // against the vendors list once it's loaded (see the useEffect below),
      // the same way ViewAdvanceForm/EditAdvanceForm do it.
      setSelectedVendorCode(item?.VendorCode || "");
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
      await getVendors();
      await getItemById(itemId);
    };
    void loadData();
  }, [context, itemId]);

  // Resolve the VendorMaster Id by matching the saved VendorCode text against the
  // vendors list (same pattern as ViewAdvanceForm/EditAdvanceForm), then use that
  // real Id to pull Previous Advances. This fixes the table always showing empty,
  // which happened because selectedVendorId was previously read from
  // item.VendorCode.Id — undefined, since VendorCode is text, not a Lookup.
  useEffect(() => {
    if (vendors.length > 0 && selectedVendorCode) {
      const match = vendors.find((v) => v.VendorCode === selectedVendorCode);
      if (match) {
        setSelectedVendorId(match.Id);
        void getPreviousAdvances(match.Id);
      }
    }
  }, [vendors, selectedVendorCode]);

  const handleApprove = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    setIsProcessing(true);
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

      const nextStep = flow[currentIndex + 1];
      let nextApproverId: number | null = null;
      let finalStatus = "Pending for Approval";

      if (nextStep) {
        nextStep.Status = "In Progress";
        nextApproverId = nextStep.Id;

        // If the next approver's Role is "Performer", set Pending for Vouching Update
        if (nextStep.Role === "Performer") {
          finalStatus = "Pending for Vouching Update";
        } else {
          finalStatus = "Pending for Approval";
        }
      } else {
        // No next step at all — fully approved
        finalStatus = "Approved";
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
          Status: finalStatus,
          ApprovalMatrix: JSON.stringify(flow),
          CurrentApproverId: nextApproverId,
          WorkflowHistory: JSON.stringify(history),
        });

      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Approved successfully.",
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
      setIsProcessing(false);
    }
  };

  const handleSendBack = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    setIsProcessing(true);
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

      flow[currentIndex].Status = "Send Back";
      let previousApproverId = null;
      if (flow[currentIndex - 1]) {
        flow[currentIndex - 1].Status = "In Progress";
        previousApproverId = flow[currentIndex - 1].Id;
      }

      const history = itemData.WorkflowHistory
        ? JSON.parse(itemData.WorkflowHistory)
        : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Send Back",
          ApprovalMatrix: JSON.stringify(flow),
          CurrentApproverId: previousApproverId,
          WorkflowHistory: JSON.stringify(history),
        });
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Send Back successfully.",
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
      setIsProcessing(false);
    }
  };

  const handleReject = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    setIsProcessing(true);
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

      flow[currentIndex].Status = "Reject";
      const history = itemData.WorkflowHistory
        ? JSON.parse(itemData.WorkflowHistory)
        : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Rejected",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Reject",
          ApprovalMatrix: JSON.stringify(flow),
          CurrentApproverId: null,
          WorkflowHistory: JSON.stringify(history),
        });
      await Swal.fire({
        icon: "success",
        title: "Success",
        text: "Rejected successfully.",
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
      setIsProcessing(false);
    }
  };

  const handleExit = () => onClose();

  const overallStatus: string = itemData?.Status || "";

  const buildRibbonSteps = () => {
    const initiatorStep = {
      Role: "Initiator",
      Name: itemData?.EmployeeName || "",
      Status: "Approved",
    };
    const approverSteps = approvalMatrix.filter((a) => a.Role !== "Initiator");
    const steps = [initiatorStep, ...approverSteps];

    if (overallStatus === "Paid") {
      return steps.map((s) => ({ ...s, _color: "approved" }));
    }

    if (overallStatus === "Reject") {
      const rejectIndex = steps.findIndex(
        (s) => s.Status === "Reject" || s.Status === "Rejected",
      );
      return steps.map((s, idx) => {
        if (rejectIndex === -1) return { ...s, _color: "" };
        if (idx === rejectIndex) return { ...s, _color: "rejected" };
        if (idx < rejectIndex) return { ...s, _color: "approved" };
        return { ...s, _color: "upcoming" };
      });
    }

    if (overallStatus === "Send Back") {
      return steps.map((s) =>
        s.Role === "Initiator"
          ? { ...s, _color: "active" }
          : { ...s, _color: "upcoming" },
      );
    }

    return steps.map((s) => {
      if (s.Status === "Approved") return { ...s, _color: "approved" };
      if (s.Status === "In Progress") return { ...s, _color: "active" };
      return { ...s, _color: "upcoming" };
    });
  };

  const getStepClass = (color: string) => {
    switch (color) {
      case "approved":
        return "approved";
      case "active":
        return "active";
      case "upcoming":
        return "upcoming";
      case "rejected":
        return "rejected";
      default:
        return "";
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
              <h1>Capex Payment Approval form </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {buildRibbonSteps().map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${getStepClass(a._color)}`}
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
                    <label className="font">Vendor Name</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.VendorName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.VendorCode}</label>
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
                    <label className="font">PO Basic Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{itemData.POBasicAmount}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO GST Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData.POGSTAmount}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Other Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{itemData.POOtherAmount}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Total PO Amount</label> :
                    &nbsp;&nbsp;
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
                    <label className="font">MRN Basic Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNBasicAmount}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN GST Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{itemData?.MRNGSTAmount}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Other Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNOtherAmount}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRNAmount including GST</label> :
                    &nbsp;&nbsp;
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

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Payment Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                        <thead
                          className="text-white"
                          style={{ backgroundColor: "rgb(60, 62, 69)" }}
                        >
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
                              <td
                                colSpan={7}
                                style={{ textAlign: "center", padding: "10px" }}
                              >
                                No previous advances available
                              </td>
                            </tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => {
                              const pending = Math.max(
                                0,
                                Number(item.RequestAdvanceAmount || 0) -
                                  Number(item.PaidAmount || 0),
                              );
                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">
                                    {item.RequestAdvanceAmount}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.Created
                                      ? new Date(
                                          item.Created,
                                        ).toLocaleDateString("en-GB")
                                      : ""}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.VoucherDate
                                      ? new Date(
                                          item.VoucherDate,
                                        ).toLocaleDateString("en-GB")
                                      : ""}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.VouchingNumber}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.PaidAmount}
                                  </td>
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

              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="heading1" style={{ marginTop: "10px" }}>
                    <label>Final Payment Details</label>
                  </div>
                  <div className="main-formcontainer">
                    <div className="row mb-20">
                      <div className="col-md-4">
                        <label className="font">Final Payment Against PO</label>{" "}
                        : &nbsp;&nbsp;
                        <label className="fonttext">
                          {itemData?.FinalPaymentAgainstPO ? "Yes" : "No"}
                        </label>
                      </div>
                    </div>
                    {itemData?.FinalPaymentAgainstPO && (
                      <div className="row mb-20"></div>
                    )}
                    {!itemData?.FinalPaymentAgainstPO && (
                      <div className="row mb-20">
                        <div className="col-md-4">
                          <label className="font">
                            Installation Request Number
                          </label>{" "}
                          : &nbsp;&nbsp;
                          <label className="fonttext">
                            {itemData?.InstallationRequestNumber}
                          </label>
                        </div>
                      </div>
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

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Approver Action</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                    <textarea
                      className="font-control"
                      value={approverRemarks}
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
                      <p>No attachments found</p>
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
                    <a className="submit-btn" onClick={handleApprove}>
                      Approve
                    </a>
                    <a className="Rework-btn" onClick={handleSendBack}>
                      Send Back
                    </a>
                    <a className="Reject-btn" onClick={handleReject}>
                      Reject
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

export default ApproverAdvanceForm;
