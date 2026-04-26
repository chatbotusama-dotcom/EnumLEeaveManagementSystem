import * as React from 'react';
import styles from './ApplyLeave.module.scss';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";

interface IFormState {
  leaveTypeId: number;
  leaveDurationType: 'Full Day' | 'Half Day';
  startDate: string;
  endDate: string;
  totalDays: number;
  halfDayType: 'First Half' | 'Second Half';
  reason: string;
  newAttachments: File[];
  otherLeaveType: string;
  existingAttachments: any[];
  originalLeaveTypeId: number;
  originalLeaveTypeWasSick: boolean;
}

interface IFormErrors {
  leaveType?: string;
  startDate?: string;
  endDate?: string;
  reason?: string;
  overlap?: string;
  attachment?: string;
  submit?: string;
  weekend?: string;
  manager?: string;
  probation?: string;
  otherLeaveType?: string;
  insufficientQuota?: string;
}

interface IEmployeeData {
  id: number;
  name: string;
  department: string;
  departmentId: number;
  country: string;
  countryId: number;
  countryCode: string;
  countryCodeId: number;
  managerId: number;
  managerName: string;
  managerEmail: string;
  joinDate: string;
  employmentType: string;
  userImageUrl: string;
}

interface ILeaveType {
  Id: number;
  Title: string;
}

interface ILeaveQuota {
  Id: number;
  CountryId: number;
  CountryCodeId: number;
  Leaves: number;
  AnnualLeaves: number;
  CasualLeaves: number;
  SickLeaves: number;
  OtherLeaves: number;
}

interface IEmployeeLeaveBalance {
  Id: number;
  LeavesBalance: number;
  Used: number;
  Remaining: number;
  Year: string;
}

interface ILeaveCycle {
  startDate: Date;
  endDate: Date;
  cycleNumber: number;
}

interface ILeaveRequest {
  Id: number;
  StartDate: string;
  EndDate: string;
  Status: string;
  TotalDays: number;
  LeaveTypeId: number;
  EmployeeYearCycle: number;
  Resubmit?: boolean;
  EmployeeId?: number;
}

interface IEditLeaveState {
  formState: IFormState;
  errors: IFormErrors;
  isSubmitting: boolean;
  showProbationPopup: boolean;
  showConfirmationPopup: boolean;
  showSuccessPopup: boolean;
  showDeleteConfirmPopup: boolean;
  deleteAttachmentId: any | null;
  weekendError: string | null;
  showHalfDayType: boolean;
  currentUser: any;
  leaveQuota: ILeaveQuota | null;
  employeeLeaveBalance: IEmployeeLeaveBalance | null;
  currentCycle: ILeaveCycle | null;
  usedLeavesInCycle: {
    AnnualLeaves: number;
    CasualLeaves: number;
    SickLeaves: number;
    OtherLeaves: number;
  };
  employee: IEmployeeData | null;
  loading: boolean;
  existingRequests: ILeaveRequest[] | null;
  requestId: number | null;
  leaveRequestData: any | null;
  isEditable: boolean;
  isResubmitMode: boolean;
  viewOnly: boolean;
  userRole: 'employee' | 'hr' | 'executive' | 'departmentManager' | 'none';
  canDownloadAttachments: boolean;
}

const STATIC_LEAVE_TYPES: ILeaveType[] = [
  { Id: 1, Title: 'Annual Leave' },
  { Id: 2, Title: 'Casual Leave' },
  { Id: 3, Title: 'Sick Leave' },
  { Id: 4, Title: 'Other Leave' }
];

const EDITABLE_STATUSES = ['Send Back by Manager', 'Send Back by HR', 'Send Back by Executive'];
const PENDING_STATUSES = ['Pending', 'Pending on Manager', 'Pending on HR', 'Pending on Executive'];

const FORM_INITIAL_STATE: IFormState = {
  leaveTypeId: 0,
  leaveDurationType: 'Full Day',
  startDate: '',
  endDate: '',
  totalDays: 0,
  halfDayType: 'First Half',
  reason: '',
  newAttachments: [],
  otherLeaveType: '',
  existingAttachments: [],
  originalLeaveTypeId: 0,
  originalLeaveTypeWasSick: false
};

const FILE_CONFIG = {
  MAX_SIZE: 5 * 1024 * 1024,
  ACCEPTED_TYPES: ['application/pdf', 'image/jpeg', 'image/png', 'image/jpg'] as const,
};

type AcceptedFileType = typeof FILE_CONFIG.ACCEPTED_TYPES[number];

const isWeekend = (date: Date): boolean => date.getDay() === 5 || date.getDay() === 6;

const calculateTotalDays = (startDate: Date, endDate: Date, isHalfDay: boolean): number => {
  if (isHalfDay) return 0.5;
  let days = 0;
  const current = new Date(startDate);
  const end = new Date(endDate);
  while (current <= end) {
    days++;
    current.setDate(current.getDate() + 1);
  }
  return days;
};

const validateWeekendDates = (startDate: Date, endDate: Date): string | null => {
  if (isWeekend(startDate)) return 'Start date cannot be a weekend (Friday or Saturday)';
  if (isWeekend(endDate)) return 'End date cannot be a weekend (Friday or Saturday)';
  return null;
};

const isDateOverlap = (start1: Date, end1: Date, start2: Date, end2: Date): boolean =>
  start1 <= end2 && end1 >= start2;

const isEmployeeEligibleForLeaveQuota = (employmentType: string): boolean => {
  const eligibleTypes = ['Permanent', 'Contract', 'Part-time'];
  return eligibleTypes.includes(employmentType);
};

const calculateCurrentLeaveCycle = (joinDate: string): ILeaveCycle => {
  const join = new Date(joinDate);
  const today = new Date();

  let cycleStart = new Date(join);
  let cycleEnd = new Date(join);
  cycleEnd.setFullYear(cycleEnd.getFullYear() + 1);
  cycleEnd.setDate(cycleEnd.getDate() - 1);
  let cycleNumber = 1;

  while (today > cycleEnd) {
    cycleStart = new Date(cycleEnd);
    cycleStart.setDate(cycleStart.getDate() + 1);
    cycleEnd = new Date(cycleStart);
    cycleEnd.setFullYear(cycleEnd.getFullYear() + 1);
    cycleEnd.setDate(cycleEnd.getDate() - 1);
    cycleNumber++;
  }

  return { startDate: cycleStart, endDate: cycleEnd, cycleNumber };
};

const redirectToCurrentPageWithoutRequestId = () => {
  const url = new URL(window.location.href);
  url.searchParams.delete('RequestID');
  window.location.href = url.toString();
};

export default class EditLeave extends React.Component<{ onLeaveSubmitted?: () => void }, IEditLeaveState> {

  constructor(props: { onLeaveSubmitted?: () => void }) {
    super(props);
    this.state = {
      formState: FORM_INITIAL_STATE,
      errors: {},
      isSubmitting: false,
      showProbationPopup: false,
      showConfirmationPopup: false,
      showSuccessPopup: false,
      showDeleteConfirmPopup: false,
      deleteAttachmentId: null,
      weekendError: null,
      showHalfDayType: false,
      currentUser: null,
      leaveQuota: null,
      employeeLeaveBalance: null,
      currentCycle: null,
      usedLeavesInCycle: {
        AnnualLeaves: 0,
        CasualLeaves: 0,
        SickLeaves: 0,
        OtherLeaves: 0
      },
      employee: null,
      loading: true,
      existingRequests: null,
      requestId: null,
      leaveRequestData: null,
      isEditable: false,
      isResubmitMode: false,
      viewOnly: false,
      userRole: 'none',
      canDownloadAttachments: false
    };
  }

  async componentDidMount() {
    await this.fetchCurrentUser();
  }

  fetchCurrentUser = async () => {
    try {
      const user = await sp.web.currentUser();
      this.setState({ currentUser: user }, () => {
        this.getRequestIdAndLoad();
      });
    } catch (err) {
      this.setState({ loading: false, errors: { submit: 'Failed to load user information' } });
    }
  };

  // ✅ Method to fetch HR and Executive department managers
  fetchDepartmentManagers = async () => {
    try {
      // Fetch Human Resources department manager
      const hrDept = await sp.web.lists
        .getByTitle('Department')
        .items
        .select('Id', 'Title', 'DepartmentManager/Id', 'DepartmentManager/Title', 'DepartmentManager/EMail')
        .expand('DepartmentManager')
        .filter(`Title eq 'Human Resources'`)
        .top(1)
        .get();

      // Fetch Executive department manager
      const executiveDept = await sp.web.lists
        .getByTitle('Department')
        .items
        .select('Id', 'Title', 'DepartmentManager/Id', 'DepartmentManager/Title', 'DepartmentManager/EMail')
        .expand('DepartmentManager')
        .filter(`Title eq 'Executive'`)
        .top(1)
        .get();

      return {
        hrManager: hrDept && hrDept.length > 0 ? hrDept[0].DepartmentManager : null,
        executiveManager: executiveDept && executiveDept.length > 0 ? executiveDept[0].DepartmentManager : null
      };
    } catch (err) {
      console.error("Error fetching department managers:", err);
      return { hrManager: null, executiveManager: null };
    }
  };

  // ✅ Check if user is Department Manager of given department
  checkIfDepartmentManager = async (departmentId: number, userEmail: string): Promise<boolean> => {
    try {
      if (!departmentId) return false;
      
      const deptData = await sp.web.lists
        .getByTitle('Department')
        .items
        .select('Id', 'Title', 'DepartmentManager/Id', 'DepartmentManager/Title', 'DepartmentManager/EMail')
        .expand('DepartmentManager')
        .filter(`Id eq ${departmentId}`)
        .top(1)
        .get();

      if (deptData && deptData.length > 0 && deptData[0].DepartmentManager) {
        return deptData[0].DepartmentManager.EMail === userEmail;
      }
      return false;
    } catch (err) {
      console.error("Error checking department manager:", err);
      return false;
    }
  };

  // ✅ Check user role (Employee, HR, Executive, or Department Manager)
  checkUserRole = async (currentUserEmail: string, leaveRequestEmployeeId: number, employeeDepartmentId?: number): Promise<'employee' | 'hr' | 'executive' | 'departmentManager' | 'none'> => {
    const { hrManager, executiveManager } = await this.fetchDepartmentManagers();
    
    // Check if user is HR Manager
    if (hrManager && hrManager.EMail === currentUserEmail) {
      return 'hr';
    }
    
    // Check if user is Executive Manager
    if (executiveManager && executiveManager.EMail === currentUserEmail) {
      return 'executive';
    }
    
    // Check if user is Department Manager of the employee
    if (employeeDepartmentId) {
      const isDeptManager = await this.checkIfDepartmentManager(employeeDepartmentId, currentUserEmail);
      if (isDeptManager) {
        return 'departmentManager';
      }
    }
    
    // Check if user is the employee who created the request
    const currentUserId = this.state.currentUser?.Id;
    if (currentUserId === leaveRequestEmployeeId) {
      return 'employee';
    }
    
    return 'none';
  };

  getRequestIdAndLoad = () => {
    const urlParams = new URLSearchParams(window.location.search);
    const requestIdParam = urlParams.get('RequestID');

    if (requestIdParam && requestIdParam.trim() !== '') {
      const requestId = parseInt(requestIdParam, 10);
      if (!isNaN(requestId) && requestId > 0) {
        this.setState({ requestId }, () => {
          this.fetchLeaveRequestDetails();
        });
      } else {
        this.setState({ loading: false, errors: { submit: 'Invalid Request ID format' } });
      }
    } else {
      this.setState({ loading: false, errors: { submit: 'No Request ID found in URL' } });
    }
  };

  // ✅ Fetch employee department ID for role check
  fetchEmployeeDepartmentId = async (employeeId: number): Promise<number> => {
    try {
      const items = await sp.web.lists
        .getByTitle('Employee Information List')
        .items
        .select('Department/Id')
        .expand('Department')
        .filter(`EmployeeName/Id eq ${employeeId}`)
        .top(1)
        .get();

      if (items && items.length > 0) {
        return items[0].Department?.Id || 0;
      }
      return 0;
    } catch (err) {
      console.error("Error fetching employee department:", err);
      return 0;
    }
  };

  // ✅ Method to fetch leave request details first, then check role
  fetchLeaveRequestDetails = async () => {
    try {
      const { requestId, currentUser } = this.state;

      const item = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId!)
        .select('Id', 'EmployeeId', 'LeaveTypeId', 'OtherLeaveType', 'LeaveDurationType',
          'HalfDayType', 'StartDate', 'EndDate', 'TotalDays', 'Reason', 'Status', 'AttachmentFiles', 'EmployeeYearCycle')
        .expand('AttachmentFiles')
        .get();

      if (!item) {
        this.setState({ loading: false, errors: { submit: 'Leave request not found' } });
        return;
      }

      // Get employee department ID for role check
      const employeeDepartmentId = await this.fetchEmployeeDepartmentId(item.EmployeeId);
      
      // Check user role based on employee ID and department
      const userRole = await this.checkUserRole(currentUser.Email, item.EmployeeId, employeeDepartmentId);
      
      let isEditable = false;
      let isResubmitMode = false;
      let viewOnly = false;
      let canDownloadAttachments = false;

      if (userRole === 'employee') {
        // Employee logic - can edit only in certain statuses
        const isEditableByStatus = EDITABLE_STATUSES.includes(item.Status);
        const isPendingStatus = PENDING_STATUSES.includes(item.Status);
        isResubmitMode = isEditableByStatus;
        viewOnly = !isEditableByStatus && !isPendingStatus;
        isEditable = isEditableByStatus;
        canDownloadAttachments = true;
        
        // Load employee data for quota calculation
        await this.fetchEmployeeData(item.EmployeeId);
      } 
      else if (userRole === 'hr' || userRole === 'executive' || userRole === 'departmentManager') {
        // HR/Executive/Department Manager logic - view only with download capability
        viewOnly = true;
        isEditable = false;
        isResubmitMode = false;
        canDownloadAttachments = true;
        
        // Load employee data for display (view only)
        await this.fetchEmployeeDataForView(item.EmployeeId);
      }
      else {
        // Unauthorized user
        viewOnly = true;
        canDownloadAttachments = false;
        this.setState({
          loading: false,
          viewOnly: true,
          errors: { submit: 'You are not authorized to view this leave request.' }
        });
        return;
      }

      // Convert UTC to local date
      const convertUTCToLocalDate = (utcDateString: string): string => {
        if (!utcDateString) return '';
        const date = new Date(utcDateString);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      };

      const startDateFormatted = convertUTCToLocalDate(item.StartDate);
      const endDateFormatted = convertUTCToLocalDate(item.EndDate);

      this.setState({
        formState: {
          leaveTypeId: item.LeaveTypeId,
          leaveDurationType: item.LeaveDurationType,
          startDate: startDateFormatted,
          endDate: endDateFormatted,
          totalDays: item.TotalDays,
          halfDayType: item.HalfDayType || '',
          reason: item.Reason,
          newAttachments: [],
          otherLeaveType: item.OtherLeaveType || '',
          existingAttachments: item.AttachmentFiles || [],
          originalLeaveTypeId: item.LeaveTypeId,
          originalLeaveTypeWasSick: item.LeaveTypeId === 3
        },
        showHalfDayType: item.LeaveDurationType === 'Half Day',
        isEditable: isEditable,
        isResubmitMode: isResubmitMode,
        viewOnly: viewOnly,
        userRole: userRole,
        canDownloadAttachments: canDownloadAttachments,
        leaveRequestData: item,
        loading: false
      });
    } catch (err: any) {
      console.error("Error in fetchLeaveRequestDetails:", err);
      this.setState({ loading: false, errors: { submit: err.message || 'Error fetching leave request' } });
    }
  };

  // ✅ Modified to accept employeeId parameter
  fetchEmployeeData = async (employeeId?: number) => {
    try {
      const { currentUser } = this.state;
      const email = currentUser?.Email;
      
      if (!email) {
        throw new Error('User email not found');
      }

      const items = await sp.web.lists
        .getByTitle('Employee Information List')
        .items
        .select(
          'Id',
          'EmployeeName/Id',
          'EmployeeName/Title',
          'EmployeeName/EMail',
          'Department/Id',
          'Department/Title',
          'Country/Id',
          'Country/Title',
          'CountryCode/Id',
          'CountryCode/Title',
          'JoiningDate',
          'EmploymentType'
        )
        .expand('EmployeeName', 'Department', 'Country', 'CountryCode')
        .filter(`EmployeeName/EMail eq '${email.replace(/'/g, "''")}'`)
        .top(1)
        .get();

      if (!items || items.length === 0) {
        throw new Error('No employee record found');
      }

      await this.processEmployeeData(items[0]);
    } catch (err) {
      this.setState({ loading: false, errors: { submit: 'Failed to load employee data' } });
    }
  };

  // ✅ Method for HR/Executive/Department Manager to view employee data
  fetchEmployeeDataForView = async (employeeId: number) => {
    try {
      const items = await sp.web.lists
        .getByTitle('Employee Information List')
        .items
        .select(
          'Id',
          'EmployeeName/Id',
          'EmployeeName/Title',
          'EmployeeName/EMail',
          'Department/Id',
          'Department/Title',
          'Country/Id',
          'Country/Title',
          'CountryCode/Id',
          'CountryCode/Title',
          'JoiningDate',
          'EmploymentType'
        )
        .expand('EmployeeName', 'Department', 'Country', 'CountryCode')
        .filter(`EmployeeName/Id eq ${employeeId}`)
        .top(1)
        .get();

      if (!items || items.length === 0) {
        throw new Error('No employee record found');
      }

      await this.processEmployeeDataForView(items[0]);
    } catch (err) {
      this.setState({ loading: false, errors: { submit: 'Failed to load employee data' } });
    }
  };

  // ✅ Method to process employee data for HR/Executive/Department Manager view
  processEmployeeDataForView = async (empData: any) => {
    const employee: IEmployeeData = {
      id: empData.EmployeeName?.Id || 0,
      name: empData.EmployeeName?.Title || '',
      department: empData.Department?.Title || '',
      departmentId: empData.Department?.Id || 0,
      country: empData.Country?.Title || '',
      countryId: empData.Country?.Id || 0,
      countryCode: empData.CountryCode?.Title || '',
      countryCodeId: empData.CountryCode?.Id || 0,
      managerId: 0,
      managerName: '',
      managerEmail: '',
      joinDate: empData.JoiningDate || '',
      employmentType: empData.EmploymentType || 'Permanent',
      userImageUrl: '',
    };

    this.setState({ employee });
  };

  fetchDepartmentData = async (departmentId: number) => {
    try {
      if (!departmentId) return null;

      const items = await sp.web.lists
        .getByTitle('Department')
        .items
        .select('Id', 'Title', 'DepartmentManager/Id', 'DepartmentManager/Title', 'DepartmentManager/EMail')
        .expand('DepartmentManager')
        .filter(`Id eq ${departmentId}`)
        .top(1)
        .get();

      return items && items.length > 0 ? items[0] : null;
    } catch (err) {
      return null;
    }
  };

  fetchLeaveQuota = async (countryId: number, countryCodeId: number) => {
    try {
      if (!countryId || !countryCodeId) {
        return null;
      }

      const items = await sp.web.lists
        .getByTitle('Leave Quota')
        .items
        .select(
          'Id',
          'Country/Id',
          'CountryCode/Id',
          'Leaves',
          'AnnualLeaves',
          'CasualLeaves',
          'SickLeaves',
          'OtherLeaves'
        )
        .expand('Country', 'CountryCode')
        .filter(`Country/Id eq ${countryId} and CountryCode/Id eq ${countryCodeId}`)
        .top(1)
        .get();

      if (items && items.length > 0) {
        const quota = items[0];
        return {
          Id: quota.Id,
          CountryId: quota.Country?.Id,
          CountryCodeId: quota.CountryCode?.Id,
          Leaves: quota.Leaves,
          AnnualLeaves: quota.AnnualLeaves,
          CasualLeaves: quota.CasualLeaves,
          SickLeaves: quota.SickLeaves,
          OtherLeaves: quota.OtherLeaves
        } as ILeaveQuota;
      }
      return null;
    } catch (err) {
      return null;
    }
  };

  fetchLeaveRequests = async (employeeId: number, cycleNumber: number) => {
    try {
      if (!employeeId) return [];

      const { requestId } = this.state;

      const filter = `EmployeeId eq ${employeeId} and (Status eq 'Pending on Manager' or Status eq 'Pending on HR' or Status eq 'Pending on Executive' or Status eq 'Approved') and EmployeeYearCycle eq ${cycleNumber} and Id ne ${requestId}`;

      const items = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .select('Id', 'StartDate', 'EndDate', 'Status', 'TotalDays', 'LeaveTypeId', 'EmployeeYearCycle')
        .filter(filter)
        .orderBy('StartDate', false)
        .get();

      return items as ILeaveRequest[];
    } catch (err) {
      return [];
    }
  };

  fetchLeaveBalance = async (employeeId: number) => {
    try {
      if (!employeeId) return null;

      const currentYear = new Date().getFullYear().toString();

      const items = await sp.web.lists
        .getByTitle('Employee Leave Balance')
        .items
        .select('Id', 'LeavesBalance', 'Used', 'Remaining', 'Year', 'EmployeeId')
        .filter(`EmployeeId eq ${employeeId} and Year eq '${currentYear}'`)
        .top(1)
        .get();

      if (items && items.length > 0) {
        return items[0] as IEmployeeLeaveBalance;
      }
      return null;
    } catch (err) {
      return null;
    }
  };

  processEmployeeData = async (empData: any) => {
    const { currentUser } = this.state;

    const departmentId = empData.Department?.Id;
    const deptData = departmentId ? await this.fetchDepartmentData(departmentId) : null;

    let managerId = 0, managerName = '', managerEmail = '';
    if (deptData?.DepartmentManager) {
      managerId = deptData.DepartmentManager.Id;
      managerName = deptData.DepartmentManager.Title;
      managerEmail = deptData.DepartmentManager.EMail;
    }

    const employee: IEmployeeData = {
      id: currentUser?.Id || 0,
      name: empData.EmployeeName?.Title || currentUser?.Title || '',
      department: empData.Department?.Title || '',
      departmentId: departmentId || 0,
      country: empData.Country?.Title || '',
      countryId: empData.Country?.Id || 0,
      countryCode: empData.CountryCode?.Title || '',
      countryCodeId: empData.CountryCode?.Id || 0,
      managerId, managerName, managerEmail,
      joinDate: empData.JoiningDate || '',
      employmentType: empData.EmploymentType || 'Permanent',
      userImageUrl: '',
    };

    this.setState({ employee }, async () => {
      const cycle = calculateCurrentLeaveCycle(employee.joinDate);
      this.setState({ currentCycle: cycle });

      await Promise.all([
        this.fetchLeaveQuota(employee.countryId, employee.countryCodeId),
        this.fetchLeaveBalance(employee.id),
        this.fetchLeaveRequests(employee.id, cycle.cycleNumber)
      ]).then(([quota, balance, requests]) => {
        this.setState({
          leaveQuota: quota,
          employeeLeaveBalance: balance,
          existingRequests: requests
        }, () => {
          this.calculateUsedLeaves();
        });
      });
    });
  };

  calculateUsedLeaves = () => {
    const { existingRequests, currentCycle } = this.state;
    if (existingRequests && existingRequests.length > 0 && currentCycle) {
      const used = { AnnualLeaves: 0, CasualLeaves: 0, SickLeaves: 0, OtherLeaves: 0 };

      existingRequests.forEach((request: ILeaveRequest) => {
        const totalDays = request.TotalDays || 0;
        if (request.LeaveTypeId === 1) used.AnnualLeaves += totalDays;
        else if (request.LeaveTypeId === 2) used.CasualLeaves += totalDays;
        else if (request.LeaveTypeId === 3) used.SickLeaves += totalDays;
        else if (request.LeaveTypeId === 4) used.OtherLeaves += totalDays;
      });

      this.setState({ usedLeavesInCycle: used });
    }
  };

  downloadAttachment = async (attachment: any) => {
    try {
      const { requestId } = this.state;
      
      const blob = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId!)
        .attachmentFiles
        .getByName(attachment.FileName)
        .getBlob();
      
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = attachment.FileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Error downloading attachment:', error);
      this.setState({ errors: { attachment: 'Failed to download attachment' } });
    }
  };

  openAttachment = async (attachment: any) => {
    try {
      const { requestId } = this.state;
      const fileName = attachment.FileName;

      // Try to get server-relative file URL first
      const fileUrl = attachment.ServerRelativeUrl || attachment.FileRef || attachment.FilePath;

      if (fileUrl) {
        // Open directly in browser / Office Online
        window.open(fileUrl, '_blank');
        return;
      }

      // Fallback: blob method
      const blob = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId!)
        .attachmentFiles
        .getByName(fileName)
        .getBlob();

      const url = window.URL.createObjectURL(blob);
      window.open(url, '_blank');
      setTimeout(() => {
        window.URL.revokeObjectURL(url);
      }, 1000);
    } catch (error) {
      console.error('Error opening attachment:', error);
      this.setState({ errors: { attachment: 'Failed to open attachment' } });
    }
  };

  isProbation = (): boolean => {
    const { employee } = this.state;
    if (!employee?.employmentType) return false;
    return employee.employmentType.toLowerCase() === 'probation';
  };

  getAvailableQuota = (leaveTypeId: number): number => {
    const { leaveQuota, usedLeavesInCycle, employeeLeaveBalance, currentCycle } = this.state;

    if (this.isProbation()) return 0;
    if (!isEmployeeEligibleForLeaveQuota(this.state.employee?.employmentType || '')) return 0;
    if (!leaveQuota) return 0;

    let baseQuota = 0;
    switch (leaveTypeId) {
      case 1:
        baseQuota = leaveQuota.AnnualLeaves;
        break;
      case 2:
        baseQuota = leaveQuota.CasualLeaves;
        break;
      case 3:
        baseQuota = leaveQuota.SickLeaves;
        break;
      case 4:
        baseQuota = leaveQuota.OtherLeaves;
        break;
      default:
        return 0;
    }

    let totalAvailable = baseQuota;

    if (leaveTypeId !== 4 && currentCycle && currentCycle.cycleNumber > 1 && employeeLeaveBalance) {
      totalAvailable = baseQuota + (employeeLeaveBalance.Remaining || 0);
    }

    let usedLeaves = 0;
    switch (leaveTypeId) {
      case 1:
        usedLeaves = usedLeavesInCycle.AnnualLeaves;
        break;
      case 2:
        usedLeaves = usedLeavesInCycle.CasualLeaves;
        break;
      case 3:
        usedLeaves = usedLeavesInCycle.SickLeaves;
        break;
      case 4:
        usedLeaves = usedLeavesInCycle.OtherLeaves;
        break;
    }

    return Math.max(0, totalAvailable - usedLeaves);
  };

  canApplyForLeave = (leaveTypeId: number, requestedDays: number): boolean => {
    if (this.isProbation()) return false;
    if (!isEmployeeEligibleForLeaveQuota(this.state.employee?.employmentType || '')) return false;

    const available = this.getAvailableQuota(leaveTypeId);
    return requestedDays <= available;
  };

  validateForm = (): IFormErrors => {
    const { formState, existingRequests } = this.state;
    const newErrors: IFormErrors = {};

    if (this.isProbation()) {
      newErrors.probation = 'You are on probation period. You cannot apply for any leave.';
      return newErrors;
    }

    if (!formState.leaveTypeId || formState.leaveTypeId === 0) newErrors.leaveType = 'Please select a leave type';
    if (formState.leaveTypeId === 4 && !formState.otherLeaveType.trim()) newErrors.otherLeaveType = 'Please specify other leave type';

    if (!formState.startDate) {
      newErrors.startDate = 'Please select start date';
    } else {
      const startDate = new Date(formState.startDate);
      if (isWeekend(startDate)) newErrors.startDate = 'Start date cannot be a weekend (Friday or Saturday)';
    }

    if (formState.leaveDurationType === 'Full Day' && !formState.endDate) {
      newErrors.endDate = 'Please select end date';
    } else if (formState.startDate && formState.endDate) {
      const start = new Date(formState.startDate);
      const end = new Date(formState.endDate);
      if (end < start) newErrors.endDate = 'End date cannot be before start date';
      if (isWeekend(end)) newErrors.endDate = 'End date cannot be a weekend (Friday or Saturday)';
    }

    if (!formState.reason) {
      newErrors.reason = 'Please provide a reason for leave';
    } else if (formState.reason.length < 10) {
      newErrors.reason = 'Please provide a more detailed reason (minimum 10 characters)';
    }

    if (existingRequests && existingRequests.length > 0 && formState.startDate && formState.endDate) {
      const newStart = new Date(formState.startDate);
      const newEnd = new Date(formState.endDate);
      const hasOverlap = existingRequests.some((leave: ILeaveRequest) => {
        const existingStart = new Date(leave.StartDate);
        const existingEnd = new Date(leave.EndDate);
        return isDateOverlap(newStart, newEnd, existingStart, existingEnd);
      });
      if (hasOverlap) newErrors.overlap = 'You already have a leave request for these dates';
    }

    if (formState.leaveTypeId && formState.totalDays > 0) {
      const canApply = this.canApplyForLeave(formState.leaveTypeId, formState.totalDays);
      const selectedLeaveType = STATIC_LEAVE_TYPES.find(type => type.Id === formState.leaveTypeId);
      const availableQuota = this.getAvailableQuota(formState.leaveTypeId);

      if (!canApply) {
        newErrors.insufficientQuota = `Insufficient ${selectedLeaveType?.Title} quota. Available: ${availableQuota} days, Requested: ${formState.totalDays} days.`;
      }
    }

    const selectedLeaveType = STATIC_LEAVE_TYPES.find(type => type.Id === formState.leaveTypeId);
    if (selectedLeaveType?.Title === 'Sick Leave' && formState.totalDays > 1 && formState.newAttachments.length === 0 && formState.existingAttachments.length === 0) {
      newErrors.attachment = 'Medical certificate required for sick leave exceeding 1 day';
    }

    for (const file of formState.newAttachments) {
      const isValidType = FILE_CONFIG.ACCEPTED_TYPES.includes(file.type as AcceptedFileType);
      if (!isValidType) newErrors.attachment = 'Only PDF, JPG, and PNG files are allowed';
      else if (file.size > FILE_CONFIG.MAX_SIZE) newErrors.attachment = 'File size must be less than 5MB';
    }

    return newErrors;
  };

  handleFieldChange = (field: keyof IFormState, value: any) => {
    if (!this.state.isEditable && !this.state.isResubmitMode) return;

    this.setState(prevState => ({
      formState: { ...prevState.formState, [field]: value },
      errors: { ...prevState.errors, [field]: undefined }
    }), () => {
      if (field === 'leaveDurationType') {
        const isHalfDay = value === 'Half Day';
        this.setState({ showHalfDayType: isHalfDay });

        if (isHalfDay && this.state.formState.startDate) {
          this.setState(prevState => ({
            formState: { ...prevState.formState, endDate: prevState.formState.startDate }
          }), () => {
            this.updateTotalDays();
          });
        } else if (!isHalfDay) {
          const { startDate, endDate } = this.state.formState;
          if (startDate && endDate && new Date(startDate) > new Date(endDate)) {
            this.setState(prevState => ({
              formState: { ...prevState.formState, endDate: '' }
            }), () => {
              this.updateTotalDays();
            });
          } else {
            this.updateTotalDays();
          }
        }
      } else if (field === 'startDate') {
        const isHalfDay = this.state.formState.leaveDurationType === 'Half Day';

        if (isHalfDay) {
          this.setState(prevState => ({
            formState: { ...prevState.formState, endDate: value }
          }), () => {
            this.updateTotalDays();
          });
        } else {
          const { endDate } = this.state.formState;
          if (value && endDate && new Date(value) > new Date(endDate)) {
            this.setState(prevState => ({
              formState: { ...prevState.formState, endDate: '' }
            }), () => {
              this.updateTotalDays();
            });
          } else {
            this.updateTotalDays();
          }
        }
      } else if (field === 'endDate') {
        this.updateTotalDays();
      } else if (field === 'leaveTypeId') {
        this.updateTotalDays();
      }
    });
  };

  handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const filesArray = Array.from(e.target.files);
      this.setState(prevState => ({
        formState: {
          ...prevState.formState,
          newAttachments: [...prevState.formState.newAttachments, ...filesArray]
        },
        errors: { ...prevState.errors, attachment: undefined }
      }));
    }
  };

  removeNewAttachment = (index: number) => {
    this.setState(prevState => ({
      formState: {
        ...prevState.formState,
        newAttachments: prevState.formState.newAttachments.filter((_, i) => i !== index)
      }
    }));
  };

  showDeleteConfirm = (attachment: any) => {
    this.setState({
      showDeleteConfirmPopup: true,
      deleteAttachmentId: attachment
    });
  };

  deleteSingleAttachment = async () => {
    const { requestId, deleteAttachmentId, formState } = this.state;

    try {
      await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId!)
        .attachmentFiles
        .getByName(deleteAttachmentId.FileName)
        .delete();

      const updatedAttachments = formState.existingAttachments.filter(
        (att: any) => att.FileName !== deleteAttachmentId.FileName
      );

      this.setState(prevState => ({
        formState: {
          ...prevState.formState,
          existingAttachments: updatedAttachments
        },
        showDeleteConfirmPopup: false,
        deleteAttachmentId: null
      }));
    } catch (error) {
      console.error('Error deleting attachment:', error);
    }
  };

  deleteAllAttachmentsIfNeeded = async () => {
    const { formState, requestId } = this.state;
    const originalWasSick = formState.originalLeaveTypeWasSick;
    const currentIsSick = formState.leaveTypeId === 3;

    if (originalWasSick && !currentIsSick && formState.existingAttachments.length > 0) {
      for (const attachment of formState.existingAttachments) {
        await sp.web.lists
          .getByTitle('Leave Requests')
          .items
          .getById(requestId!)
          .attachmentFiles
          .getByName(attachment.FileName)
          .delete();
      }
      this.setState(prevState => ({
        formState: {
          ...prevState.formState,
          existingAttachments: []
        }
      }));
    }
  };

  uploadNewAttachments = async () => {
    const { requestId, formState } = this.state;

    for (const file of formState.newAttachments) {
      const buffer = await file.arrayBuffer();
      await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId!)
        .attachmentFiles
        .add(file.name, buffer);
    }
  };

  updateTotalDays = () => {
    const { startDate, endDate, leaveDurationType } = this.state.formState;

    if (leaveDurationType === 'Half Day' && startDate) {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0.5 }
      }));
      return;
    }

    if (!startDate || !endDate) {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0 }
      }));
      return;
    }

    const start = new Date(startDate);
    const end = new Date(endDate);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return;

    const weekendErrorMsg = validateWeekendDates(start, end);
    this.setState({ weekendError: weekendErrorMsg });

    if (!weekendErrorMsg && end >= start) {
      const days = calculateTotalDays(start, end, false);
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: days }
      }));
    } else {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0 }
      }));
    }
  };

  handleSuccessClose = () => {
    this.setState({ showSuccessPopup: false });
    redirectToCurrentPageWithoutRequestId();
  };

  handleSubmit = () => {
    const { formState } = this.state;

    if (this.isProbation()) {
      this.setState({ showProbationPopup: true });
      return;
    }

    if (formState.startDate && formState.endDate) {
      const start = new Date(formState.startDate);
      const end = new Date(formState.endDate);
      const weekendErrorMsg = validateWeekendDates(start, end);
      if (weekendErrorMsg) {
        this.setState({ errors: { weekend: weekendErrorMsg } });
        return;
      }
    }

    const validationErrors = this.validateForm();
    if (Object.keys(validationErrors).length > 0) {
      this.setState({ errors: validationErrors });
      return;
    }

    if (!this.state.employee?.managerId) {
      this.setState({ errors: { manager: 'No manager assigned. Please contact HR.' } });
      return;
    }

    this.setState({ showConfirmationPopup: true });
  };

  redirectToDashboard = () => {
    this.setState({ showProbationPopup: false });
    redirectToCurrentPageWithoutRequestId();
  };

  getNewStatusAndClearComments = (originalStatus: string): { 
    newStatus: string; 
    clearManagerComment: boolean; 
    clearHRComment: boolean; 
    clearExecutiveComment: boolean;
    isResubmit: boolean;
  } => {
    if (originalStatus === 'Send Back by Manager') {
      return {
        newStatus: 'Pending on Manager',
        clearManagerComment: true,
        clearHRComment: false,
        clearExecutiveComment: false,
        isResubmit: true
      };
    } else if (originalStatus === 'Send Back by HR') {
      return {
        newStatus: 'Pending on HR',
        clearManagerComment: false,
        clearHRComment: true,
        clearExecutiveComment: false,
        isResubmit: true
      };
    } else if (originalStatus === 'Send Back by Executive') {
      return {
        newStatus: 'Pending on Executive',
        clearManagerComment: false,
        clearHRComment: false,
        clearExecutiveComment: true,
        isResubmit: true
      };
    }
    return {
      newStatus: 'Pending on Manager',
      clearManagerComment: false,
      clearHRComment: false,
      clearExecutiveComment: false,
      isResubmit: false
    };
  };

  confirmSubmit = async () => {
    const { formState, requestId, leaveRequestData } = this.state;

    this.setState({ isSubmitting: true, showConfirmationPopup: false });

    try {
      await this.deleteAllAttachmentsIfNeeded();

      if (formState.newAttachments.length > 0) {
        await this.uploadNewAttachments();
      }

      const { newStatus, clearManagerComment, clearHRComment, clearExecutiveComment, isResubmit } = this.getNewStatusAndClearComments(leaveRequestData?.Status || 'Send Back by Manager');

      const updateData: any = {
        LeaveTypeId: formState.leaveTypeId,
        OtherLeaveType: formState.otherLeaveType || null,
        LeaveDurationType: formState.leaveDurationType,
        HalfDayType: formState.halfDayType,
        StartDate: formState.startDate,
        EndDate: formState.endDate,
        TotalDays: formState.totalDays,
        Reason: formState.reason,
        Status: newStatus,
        Resubmit: isResubmit
      };

      if (clearManagerComment) {
        updateData.LineManagerComments = '';
      }
      if (clearHRComment) {
        updateData.HRComments = '';
      }
      if (clearExecutiveComment) {
        updateData.ExecutiveComments = '';
      }

      await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId!)
        .update(updateData);

      this.setState({
        formState: { ...this.state.formState, newAttachments: [] },
        showSuccessPopup: true
      });

    } catch (error: any) {
      this.setState({ errors: { submit: error.message || 'Error updating leave request' } });
    } finally {
      this.setState({ isSubmitting: false });
    }
  };

  componentDidUpdate(prevProps: any, prevState: IEditLeaveState) {
    if (prevState.formState.leaveDurationType !== this.state.formState.leaveDurationType &&
      this.state.formState.leaveDurationType === 'Half Day' &&
      this.state.formState.startDate) {
      this.handleFieldChange('endDate', this.state.formState.startDate);
    }
  }

  render() {
    const { loading, employee, formState, errors, isSubmitting, showProbationPopup,
      showConfirmationPopup, showSuccessPopup, showDeleteConfirmPopup, weekendError, showHalfDayType,
      isEditable, isResubmitMode, viewOnly, userRole, canDownloadAttachments } = this.state;

    if (loading) {
      return (
        <div className={styles.loadingContainer}>
          <div className={styles.spinner}></div>
          <p>Loading leave request...</p>
        </div>
      );
    }

    if (!employee) {
      return (
        <div className={styles.errorContainer}>
          <p>No employee data found. Please contact HR.</p>
          <button onClick={() => window.location.reload()}>Retry</button>
        </div>
      );
    }

    const isFormDisabled = (!isEditable && !isResubmitMode) || this.isProbation();
    const showAttachmentSection = formState.leaveTypeId === 3 && formState.totalDays > 1;
    const allAttachments = [...formState.existingAttachments, ...formState.newAttachments];

    return (
      <>
        {/* Delete Confirmation Popup */}
        {showDeleteConfirmPopup && (
          <div className={styles.probationPopupOverlay}>
            <div className={styles.probationPopup}>
              <div className={styles.probationPopupHeader}>
                <h3>Delete Attachment</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p>Are you sure you want to delete this medical evidence?</p>
              </div>
              <div className={styles.probationPopupFooter} style={{ gap: '12px' }}>
                <button
                  className={styles.probationPopupButton}
                  style={{ background: '#94a3b8' }}
                  onClick={() => this.setState({ showDeleteConfirmPopup: false, deleteAttachmentId: null })}
                >
                  Cancel
                </button>
                <button
                  className={styles.probationPopupButton}
                  onClick={this.deleteSingleAttachment}
                >
                  Delete
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Probation Popup */}
        {showProbationPopup && (
          <div className={styles.probationPopupOverlay}>
            <div className={styles.probationPopup}>
              <div className={styles.probationPopupHeader}>
                <h3>Probation Period Notice</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p>Dear <strong>{employee.name}</strong>,</p>
                <p>As per company policy, employees on <strong>Probation Period</strong> are <strong>not eligible</strong> to apply for any type of leave.</p>
                <p>Leave benefits will be available after successful completion of probation period.</p>
              </div>
              <div className={styles.probationPopupFooter}>
                <button className={styles.probationPopupButton} onClick={this.redirectToDashboard}>
                  Got It
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Confirmation Popup */}
        {showConfirmationPopup && (
          <div className={styles.probationPopupOverlay} onClick={() => this.setState({ showConfirmationPopup: false })}>
            <div className={styles.probationPopup} onClick={e => e.stopPropagation()}>
              <div className={styles.probationPopupHeader}>
                <h3>Confirm Leave Request</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p>Are you sure you want to {isResubmitMode ? 'resubmit' : 'update'} this leave request?</p>
                <p><strong>Leave Type:</strong> {STATIC_LEAVE_TYPES.find(t => t.Id === formState.leaveTypeId)?.Title}</p>
                <p><strong>Duration:</strong> {formState.totalDays} day(s)</p>
                <p><strong>Dates:</strong> {formState.startDate} to {formState.endDate}</p>
                <p><strong>Reason:</strong> {formState.reason}</p>
              </div>
              <div className={styles.probationPopupFooter} style={{ gap: '12px' }}>
                <button className={styles.probationPopupButton} style={{ background: '#94a3b8' }} onClick={() => this.setState({ showConfirmationPopup: false })}>Cancel</button>
                <button className={styles.probationPopupButton} onClick={this.confirmSubmit} disabled={isSubmitting}>
                  {isSubmitting ? 'Processing...' : 'Confirm'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Success Popup */}
        {showSuccessPopup && (
          <div className={styles.probationPopupOverlay}>
            <div className={styles.probationPopup} onClick={e => e.stopPropagation()}>
              <div className={styles.probationPopupHeader} style={{ background: '#10b981' }}>
                <h3 style={{ color: 'white' }}>✓ Success!</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p style={{ fontSize: '16px', textAlign: 'center' }}>
                  <strong>Your leave request has been {isResubmitMode ? 'resubmitted' : 'updated'} successfully!</strong>
                </p>
              </div>
              <div className={styles.probationPopupFooter}>
                <button className={styles.probationPopupButton} style={{ background: '#10b981' }} onClick={this.handleSuccessClose}>
                  OK
                </button>
              </div>
            </div>
          </div>
        )}

        <div className={styles.applyLeaveFormContent}>          
          {(Object.values(errors).some(error => error) || weekendError) && (
            <div className={styles.errorSummary}>
              {weekendError && <div className={styles.errorMessage}>{weekendError}</div>}
              {Object.values(errors).map((error, idx) => error && <div key={idx} className={styles.errorMessage}>{error}</div>)}
            </div>
          )}

          <div className={styles.applyLeaveForm}>
            {/* View Only Notice for HR/Executive/Department Manager */}
            {viewOnly && userRole !== 'employee' && (
              <div className={styles.viewOnlyNotice}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                  <circle cx="12" cy="12" r="3"></circle>
                </svg>
                <span>
                  You are viewing this leave request in <strong>read-only mode</strong>.
                  </span>
              </div>
            )}

            {/* Employee Info Fields */}
            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Employee Name</label>
                <input type="text" value={employee.name} disabled className={styles.applyLeaveDisabledField} />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Department</label>
                <input type="text" value={employee.department} disabled className={styles.applyLeaveDisabledField} />
              </div>
            </div>

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Country</label>
                <input type="text" value={employee.country} disabled className={styles.applyLeaveDisabledField} />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Manager</label>
                <input type="text" value={employee.managerName || 'Not Assigned'} disabled className={styles.applyLeaveDisabledField} />
              </div>
            </div>

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Leave Type <span className={styles.applyLeaveRequired}>*</span></label>
                <select
                  value={formState.leaveTypeId}
                  onChange={e => this.handleFieldChange('leaveTypeId', parseInt(e.target.value))}
                  className={styles.applyLeaveSelect}
                  disabled={isFormDisabled}
                >
                  <option value="0">Select Leave Type</option>
                  {STATIC_LEAVE_TYPES.map(type => (
                    <option key={type.Id} value={type.Id}>
                      {type.Title}
                    </option>
                  ))}
                </select>
                {this.isProbation() && (
                  <div className={styles.probationNote}>⚠️ You cannot apply for leave during probation period</div>
                )}
              </div>

              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Leave Duration <span className={styles.applyLeaveRequired}>*</span></label>
                <select
                  value={formState.leaveDurationType}
                  onChange={e => this.handleFieldChange('leaveDurationType', e.target.value as 'Full Day' | 'Half Day')}
                  className={styles.applyLeaveSelect}
                  disabled={isFormDisabled}
                >
                  <option value="Full Day">Full Day</option>
                  <option value="Half Day">Half Day</option>
                </select>
              </div>
            </div>

            {formState.leaveTypeId === 4 && (
              <div className={styles.applyLeaveRow}>
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>Other Leave Type <span className={styles.applyLeaveRequired}>*</span></label>
                  <input
                    type="text"
                    value={formState.otherLeaveType}
                    onChange={e => this.handleFieldChange('otherLeaveType', e.target.value)}
                    placeholder="e.g., Emergency Leave, Marriage Leave, etc."
                    className={styles.applyLeaveInput}
                    disabled={isFormDisabled}
                  />
                </div>
                {showHalfDayType && (
                  <div className={styles.applyLeaveFormGroup}>
                    <label className={styles.applyLeaveLabel}>Half Day Type <span className={styles.applyLeaveRequired}>*</span></label>
                    <select
                      value={formState.halfDayType}
                      onChange={e => this.handleFieldChange('halfDayType', e.target.value as 'First Half' | 'Second Half')}
                      className={styles.applyLeaveSelect}
                      disabled={isFormDisabled}
                    >
                      <option value="First Half">First Half</option>
                      <option value="Second Half">Second Half</option>
                    </select>
                  </div>
                )}
              </div>
            )}

            {formState.leaveTypeId !== 4 && showHalfDayType && (
              <div className={styles.applyLeaveRow}>
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>Half Day Type <span className={styles.applyLeaveRequired}>*</span></label>
                  <select
                    value={formState.halfDayType}
                    onChange={e => this.handleFieldChange('halfDayType', e.target.value as 'First Half' | 'Second Half')}
                    className={styles.applyLeaveSelect}
                    disabled={isFormDisabled}
                  >
                    <option value="First Half">First Half</option>
                    <option value="Second Half">Second Half</option>
                  </select>
                </div>
              </div>
            )}

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Start Date <span className={styles.applyLeaveRequired}>*</span></label>
                <input
                  type="date"
                  value={formState.startDate}
                  onChange={e => this.handleFieldChange('startDate', e.target.value)}
                  className={styles.applyLeaveInput}
                  disabled={isFormDisabled}
                />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>End Date {formState.leaveDurationType !== 'Half Day' && <span className={styles.applyLeaveRequired}>*</span>}</label>
                <input
                  type="date"
                  value={formState.endDate}
                  disabled={formState.leaveDurationType === 'Half Day' || isFormDisabled}
                  onChange={e => this.handleFieldChange('endDate', e.target.value)}
                  className={styles.applyLeaveInput}
                />
              </div>
            </div>

            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Total Days</label>
              <input type="number" value={formState.totalDays} disabled className={styles.applyLeaveTotalDaysField} step="0.5" />
            </div>

            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Reason for Leave <span className={styles.applyLeaveRequired}>*</span></label>
              <textarea
                value={formState.reason}
                onChange={e => this.handleFieldChange('reason', e.target.value)}
                placeholder="Please provide reason for your leave"
                rows={4}
                className={styles.applyLeaveTextarea}
                disabled={isFormDisabled}
              />
            </div>

            {/* Attachments Section - Only for Sick Leave */}
            {showAttachmentSection && (
              <>
                {allAttachments.length > 0 && (
                  <div className={styles.applyLeaveFormGroup}>
                    <label className={styles.applyLeaveLabel}>Medical Evidence</label>
                    <div className={styles.attachmentsGallery}>
                      {formState.existingAttachments.map((att: any, index: number) => (
                        <div key={`existing-${index}`} className={styles.applyLeaveFilePreview}>
                          {/* File icon based on type */}
                          {att.FileName.toLowerCase().endsWith('.pdf') ? (
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                              <path d="M4 4H20V20H4V4Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                              <path d="M8 7H16M8 11H16M8 15H13" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                            </svg>
                          ) : (
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                              <path d="M21 15V19C21 19.5304 20.7893 20.0391 20.4142 20.4142C20.0391 20.7893 19.5304 21 19 21H5C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V15M7 10L12 15M12 15L17 10M12 15V3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                            </svg>
                          )}
                          
                          {/* Clickable filename for opening */}
                          <span 
                            className={styles.applyLeaveFileName} 
                            onClick={() => this.openAttachment(att)}
                            style={{ cursor: 'pointer', textDecoration: 'underline', color: '#3b82f6' }}
                            title="Click to open"
                          >
                            {att.FileName}
                          </span>
                          
                          {/* View/Open button */}
                          {canDownloadAttachments && (
                            <span 
                              className={styles.applyLeaveViewFile} 
                              onClick={() => this.openAttachment(att)} 
                              title="Open in browser"
                            >
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
                                <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                                <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="1.5"/>
                              </svg>
                            </span>
                          )}
                          
                          {/* Download button */}
                          {canDownloadAttachments && (
                            <span 
                              className={styles.applyLeaveDownloadFile} 
                              onClick={() => this.downloadAttachment(att)} 
                              title="Download"
                            >
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
                                <path d="M21 15V19C21 19.5304 20.7893 20.0391 20.4142 20.4142C20.0391 20.7893 19.5304 21 19 21H5C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V15M7 10L12 15M12 15L17 10M12 15V3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                              </svg>
                            </span>
                          )}
                          
                          {/* Delete button (only for editable mode) */}
                          {!isFormDisabled && (
                            <span className={styles.applyLeaveRemoveFile} onClick={() => this.showDeleteConfirm(att)}>
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
                                <path d="M18 6L6 18M6 6L18 18" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                              </svg>
                            </span>
                          )}
                        </div>
                      ))}

                      {/* New attachments preview */}
                      {formState.newAttachments.map((file: File, index: number) => (
                        <div key={`new-${index}`} className={styles.applyLeaveFilePreview}>
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                            <path d="M21 15V19C21 19.5304 20.7893 20.0391 20.4142 20.4142C20.0391 20.7893 19.5304 21 19 21H5C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V15M7 10L12 15M12 15L17 10M12 15V3" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                          </svg>
                          <span>{file.name} (New)</span>
                          {!isFormDisabled && (
                            <span className={styles.applyLeaveRemoveFile} onClick={() => this.removeNewAttachment(index)}>
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
                                <path d="M18 6L6 18M6 6L18 18" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
                              </svg>
                            </span>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {!viewOnly && (
                  <div className={styles.applyLeaveFormGroup}>
                    <label className={styles.applyLeaveLabel}>
                      Upload Medical Evidence {allAttachments.length === 0 && <span className={styles.applyLeaveRequired}>*</span>}
                    </label>
                    <div
                      className={styles.applyLeaveFileDropZone}
                      onClick={() => !isFormDisabled && document.getElementById('editFileInput')?.click()}
                      style={{ cursor: isFormDisabled ? 'not-allowed' : 'pointer', opacity: isFormDisabled ? 0.6 : 1 }}
                    >
                      <div className={styles.applyLeaveUploadPrompt}>
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none">
                          <path d="M12 16V4M12 4L8 8M12 4L16 8M20 16V19C20 19.5304 19.7893 20.0391 19.4142 20.4142C19.0391 20.7893 18.5304 21 18 21H6C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V16" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                        <span>Click to upload or drag & drop</span>
                        <span className={styles.applyLeaveFileHint}>Supported: PDF, JPG, PNG (Max 5MB each)</span>
                      </div>
                    </div>
                    <input
                      id="editFileInput"
                      type="file"
                      style={{ display: 'none' }}
                      accept=".pdf,.jpg,.jpeg,.png"
                      multiple
                      onChange={this.handleFileSelect}
                      disabled={isFormDisabled}
                    />
                  </div>
                )}

                {viewOnly && canDownloadAttachments && allAttachments.length === 0 && (
                  <div className={styles.applyLeaveFormGroup}>
                    <label className={styles.applyLeaveLabel}>Medical Evidence</label>
                    <div className={styles.noAttachmentsNotice}>
                      <p>No medical evidence attached</p>
                    </div>
                  </div>
                )}
              </>
            )}

            <div className={styles.applyLeaveButtonContainer}>
              {(isEditable || isResubmitMode) && (
                <button type="button" className={styles.applyLeaveSubmitButton} onClick={this.handleSubmit} disabled={isSubmitting || this.isProbation()}>
                  {isSubmitting ? 'Processing...' : (isResubmitMode ? 'Resubmit' : 'Update Leave Request')}
                </button>
              )}
              {(viewOnly || (!isEditable && !isResubmitMode)) && (
                <button
                  type="button"
                  className={styles.applyLeaveSubmitButton}
                  onClick={redirectToCurrentPageWithoutRequestId}
                >
                  Go Back
                </button>
              )}
            </div>
          </div>
        </div>
      </>
    );
  }
}