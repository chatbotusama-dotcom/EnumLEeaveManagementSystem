import * as React from 'react';
import styles from './ExtraApplyLeave.module.scss';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";

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

interface IFormState {
  leaveTypeId: number;
  leaveDurationType: 'Full Day' | 'Half Day';
  startDate: string;
  endDate: string;
  totalDays: number;
  halfDayType: 'First Half' | 'Second Half';
  reason: string;
  attachment: File | null;
  otherLeaveType: string;
  existingAttachments: any[];
}

interface IFormErrors {
  leaveType?: string;
  startDate?: string;
  endDate?: string;
  reason?: string;
  submit?: string;
}

interface IEditLeaveState {
  formState: IFormState;
  errors: IFormErrors;
  loading: boolean;
  isSubmitting: boolean;
  requestId: number | null;
  canEdit: boolean;
  errorMessage: string | null;
  showHalfDayType: boolean;
  currentUser: any;
  employee: IEmployeeData | null;
  leaveQuota: ILeaveQuota | null;
  employeeLeaveBalance: IEmployeeLeaveBalance | null;
  currentCycle: ILeaveCycle | null;
  usedLeavesInCycle: {
    AnnualLeaves: number;
    CasualLeaves: number;
    SickLeaves: number;
    OtherLeaves: number;
  };
  showAttachmentUpload: boolean;
}

const STATIC_LEAVE_TYPES: ILeaveType[] = [
  { Id: 1, Title: 'Annual Leave' },
  { Id: 2, Title: 'Casual Leave' },
  { Id: 3, Title: 'Sick Leave' },
  { Id: 4, Title: 'Other Leave' }
];

const EDITABLE_STATUSES = ['Send Back by Manager', 'Send Back by HR', 'Send Back by Executive'];

const formatDate = (date: Date): string => date.toISOString().split('T')[0];

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

export default class EditLeave extends React.Component<{}, IEditLeaveState> {

  constructor(props: {}) {
    super(props);
    this.state = {
      formState: {
        leaveTypeId: 0,
        leaveDurationType: 'Full Day',
        startDate: '',
        endDate: '',
        totalDays: 0,
        halfDayType: 'First Half',
        reason: '',
        attachment: null,
        otherLeaveType: '',
        existingAttachments: []
      },
      errors: {},
      loading: true,
      isSubmitting: false,
      requestId: null,
      canEdit: false,
      errorMessage: null,
      showHalfDayType: false,
      currentUser: null,
      employee: null,
      leaveQuota: null,
      employeeLeaveBalance: null,
      currentCycle: null,
      usedLeavesInCycle: {
        AnnualLeaves: 0,
        CasualLeaves: 0,
        SickLeaves: 0,
        OtherLeaves: 0
      },
      showAttachmentUpload: false
    };
  }

  async componentDidMount() {
    await this.getCurrentUser();
  }

  getCurrentUser = async () => {
    try {
      const user = await sp.web.currentUser();
      this.setState({ currentUser: user }, () => {
        this.getRequestIdAndLoad();
      });
    } catch (err) {
      this.setState({ loading: false, errorMessage: 'Failed to load user information' });
    }
  };

  getRequestIdAndLoad = () => {
    const urlParams = new URLSearchParams(window.location.search);
    const requestIdParam = urlParams.get('RequestID');
    
    if (requestIdParam && requestIdParam.trim() !== '') {
      const requestId = parseInt(requestIdParam, 10);
      if (!isNaN(requestId) && requestId > 0) {
        this.setState({ requestId }, () => {
          this.loadEmployeeData();
        });
      } else {
        this.setState({ loading: false, errorMessage: 'Invalid Request ID' });
      }
    } else {
      this.setState({ loading: false, errorMessage: 'No Request ID found in URL' });
    }
  };

  loadEmployeeData = async () => {
    try {
      const { currentUser } = this.state;
      if (!currentUser?.Email) {
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
        .filter(`EmployeeName/EMail eq '${currentUser.Email.replace(/'/g, "''")}'`)
        .top(1)
        .get();

      if (!items || items.length === 0) {
        throw new Error('No employee record found');
      }

      await this.processEmployeeData(items[0]);
    } catch (err) {
      this.setState({ loading: false, errorMessage: 'Failed to load employee data' });
    }
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

      await this.fetchLeaveQuota(employee.countryId, employee.countryCodeId);
      await this.fetchLeaveBalance(employee.id);
      await this.fetchLeaveRequests(employee.id, cycle.startDate, cycle.endDate);
      await this.fetchLeaveRequestToEdit();
    });
  };

  fetchLeaveQuota = async (countryId: number, countryCodeId: number) => {
    try {
      if (!countryId || !countryCodeId) return;

      const items = await sp.web.lists
        .getByTitle('Leave Quota')
        .items
        .select('Id', 'AnnualLeaves', 'CasualLeaves', 'SickLeaves', 'OtherLeaves')
        .filter(`Country/Id eq ${countryId} and CountryCode/Id eq ${countryCodeId}`)
        .top(1)
        .get();

      if (items && items.length > 0) {
        this.setState({
          leaveQuota: {
            Id: items[0].Id,
            CountryId: countryId,
            CountryCodeId: countryCodeId,
            Leaves: items[0].Leaves,
            AnnualLeaves: items[0].AnnualLeaves,
            CasualLeaves: items[0].CasualLeaves,
            SickLeaves: items[0].SickLeaves,
            OtherLeaves: items[0].OtherLeaves
          }
        });
      }
    } catch (err) {
      console.error('Error fetching leave quota:', err);
    }
  };

  fetchLeaveBalance = async (employeeId: number) => {
    try {
      const currentYear = new Date().getFullYear().toString();
      const items = await sp.web.lists
        .getByTitle('Employee Leave Balance')
        .items
        .select('Id', 'LeavesBalance', 'Used', 'Remaining', 'Year')
        .filter(`EmployeeId eq ${employeeId} and Year eq '${currentYear}'`)
        .top(1)
        .get();

      if (items && items.length > 0) {
        this.setState({ employeeLeaveBalance: items[0] });
      }
    } catch (err) {
      console.error('Error fetching leave balance:', err);
    }
  };

  fetchLeaveRequests = async (employeeId: number, cycleStart: Date, cycleEnd: Date) => {
    try {
      const startDateStr = formatDate(cycleStart);
      const endDateStr = formatDate(cycleEnd);
      const filter = `EmployeeId eq ${employeeId} and (Status eq 'Pending on Manager' or Status eq 'Pending on HR' or Status eq 'Approved') and StartDate ge '${startDateStr}' and EndDate le '${endDateStr}'`;

      const items = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .select('Id', 'StartDate', 'EndDate', 'Status', 'TotalDays', 'LeaveTypeId')
        .filter(filter)
        .get();

      const used = { AnnualLeaves: 0, CasualLeaves: 0, SickLeaves: 0, OtherLeaves: 0 };
      items.forEach((request: any) => {
        if (request.LeaveTypeId === 1) used.AnnualLeaves += request.TotalDays;
        else if (request.LeaveTypeId === 2) used.CasualLeaves += request.TotalDays;
        else if (request.LeaveTypeId === 3) used.SickLeaves += request.TotalDays;
        else if (request.LeaveTypeId === 4) used.OtherLeaves += request.TotalDays;
      });

      this.setState({ usedLeavesInCycle: used });
    } catch (err) {
      console.error('Error fetching leave requests:', err);
    }
  };

  fetchLeaveRequestToEdit = async () => {
    try {
      const item = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(this.state.requestId!)
        .select('Id', 'EmployeeId', 'LeaveTypeId', 'OtherLeaveType', 'LeaveDurationType',
                'HalfDayType', 'StartDate', 'EndDate', 'TotalDays', 'Reason', 'Status', 'AttachmentFiles')
        .expand('AttachmentFiles')
        .get();

      if (!item) {
        this.setState({ loading: false, errorMessage: 'Leave request not found' });
        return;
      }

      // Check if current user is the employee who created this request
      const isOwner = item.EmployeeId === this.state.currentUser?.Id;
      
      // Check if status allows editing
      const canEditByStatus = EDITABLE_STATUSES.includes(item.Status);
      
      // User can edit only if they are the owner AND status is editable
      const canEdit = isOwner && canEditByStatus;

      // Remove this request from used leaves calculation
      const { usedLeavesInCycle } = this.state;
      if (item.LeaveTypeId === 1) usedLeavesInCycle.AnnualLeaves -= item.TotalDays;
      else if (item.LeaveTypeId === 2) usedLeavesInCycle.CasualLeaves -= item.TotalDays;
      else if (item.LeaveTypeId === 3) usedLeavesInCycle.SickLeaves -= item.TotalDays;
      else if (item.LeaveTypeId === 4) usedLeavesInCycle.OtherLeaves -= item.TotalDays;
      
      this.setState({ usedLeavesInCycle });

      // Check if attachment upload should be shown (Sick Leave > 1 day)
      const showAttachmentUpload = item.LeaveTypeId === 3 && item.TotalDays > 1;

      this.setState({
        formState: {
          leaveTypeId: item.LeaveTypeId,
          leaveDurationType: item.LeaveDurationType,
          startDate: item.StartDate ? item.StartDate.split('T')[0] : '',
          endDate: item.EndDate ? item.EndDate.split('T')[0] : '',
          totalDays: item.TotalDays,
          halfDayType: item.HalfDayType || 'First Half',
          reason: item.Reason,
          attachment: null,
          otherLeaveType: item.OtherLeaveType || '',
          existingAttachments: item.AttachmentFiles || []
        },
        showHalfDayType: item.LeaveDurationType === 'Half Day',
        showAttachmentUpload: showAttachmentUpload,
        canEdit: canEdit,
        loading: false
      });

      if (!isOwner) {
        this.setState({
          errorMessage: 'You are not authorized to edit this leave request. Only the employee who submitted this request can edit it.'
        });
      } else if (!canEditByStatus) {
        this.setState({
          errorMessage: `Cannot edit: Request status is "${item.Status}". Only requests with status "Send Back by Manager", "Send Back by HR", or "Send Back by Executive" can be edited.`
        });
      }
    } catch (err: any) {
      this.setState({ loading: false, errorMessage: err.message || 'Error fetching leave request' });
    }
  };

  getAvailableQuota = (leaveTypeId: number): number => {
    const { leaveQuota, usedLeavesInCycle } = this.state;
    if (!leaveQuota) return 0;

    switch (leaveTypeId) {
      case 1:
        return Math.max(0, leaveQuota.AnnualLeaves - usedLeavesInCycle.AnnualLeaves);
      case 2:
        return Math.max(0, leaveQuota.CasualLeaves - usedLeavesInCycle.CasualLeaves);
      case 3:
        return Math.max(0, leaveQuota.SickLeaves - usedLeavesInCycle.SickLeaves);
      case 4:
        return Math.max(0, leaveQuota.OtherLeaves - usedLeavesInCycle.OtherLeaves);
      default:
        return 0;
    }
  };

  handleFieldChange = (field: keyof IFormState, value: any) => {
    if (!this.state.canEdit) return;
    
    this.setState(prevState => ({
      formState: { ...prevState.formState, [field]: value },
      errors: {} // Clear all errors when user makes changes
    }), () => {
      if (field === 'leaveDurationType') {
        const isHalfDay = value === 'Half Day';
        this.setState({ showHalfDayType: isHalfDay });
        if (isHalfDay && this.state.formState.startDate) {
          this.setState(prevState => ({
            formState: { ...prevState.formState, endDate: prevState.formState.startDate }
          }), () => this.updateTotalDays());
        } else {
          this.updateTotalDays();
        }
      } else if (field === 'startDate' && this.state.formState.leaveDurationType === 'Half Day') {
        this.setState(prevState => ({
          formState: { ...prevState.formState, endDate: value }
        }), () => this.updateTotalDays());
      } else if (field === 'startDate' || field === 'endDate') {
        this.updateTotalDays();
      } else if (field === 'leaveTypeId') {
        this.updateTotalDays();
        // Update attachment upload visibility based on leave type
        const newLeaveTypeId = value as number;
        const { totalDays } = this.state.formState;
        this.setState({
          showAttachmentUpload: newLeaveTypeId === 3 && totalDays > 1
        });
      }
    });
  };

  updateTotalDays = () => {
    const { startDate, endDate, leaveDurationType } = this.state.formState;

    if (leaveDurationType === 'Half Day' && startDate) {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0.5 }
      }), () => {
        // Update attachment upload visibility after total days update
        const { leaveTypeId, totalDays } = this.state.formState;
        this.setState({
          showAttachmentUpload: leaveTypeId === 3 && totalDays > 1
        });
      });
      return;
    }

    if (!startDate || !endDate) return;

    const start = new Date(startDate);
    const end = new Date(endDate);
    
    if (end >= start) {
      const days = calculateTotalDays(start, end, false);
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: days }
      }), () => {
        // Update attachment upload visibility after total days update
        const { leaveTypeId, totalDays } = this.state.formState;
        this.setState({
          showAttachmentUpload: leaveTypeId === 3 && totalDays > 1
        });
      });
    }
  };

  handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!this.state.canEdit) return;
    if (e.target.files && e.target.files[0]) {
      this.handleFieldChange('attachment', e.target.files[0]);
    }
  };

  validateForm = (): IFormErrors => {
    const { formState } = this.state;
    const errors: IFormErrors = {};

    if (!formState.leaveTypeId || formState.leaveTypeId === 0) {
      errors.leaveType = 'Please select a leave type';
    }
    
    if (formState.leaveTypeId === 4 && !formState.otherLeaveType.trim()) {
      errors.leaveType = 'Please specify other leave type';
    }

    if (!formState.startDate) {
      errors.startDate = 'Please select start date';
    }

    if (formState.leaveDurationType === 'Full Day' && !formState.endDate) {
      errors.endDate = 'Please select end date';
    } else if (formState.startDate && formState.endDate) {
      const start = new Date(formState.startDate);
      const end = new Date(formState.endDate);
      if (end < start) {
        errors.endDate = 'End date cannot be before start date';
      }
    }

    if (!formState.reason) {
      errors.reason = 'Please provide a reason for leave';
    } else if (formState.reason.length < 10) {
      errors.reason = 'Please provide a more detailed reason (minimum 10 characters)';
    }

    // Check quota for leave types other than Other Leave
    if (formState.leaveTypeId && formState.leaveTypeId !== 4 && formState.totalDays > 0) {
      const available = this.getAvailableQuota(formState.leaveTypeId);
      if (formState.totalDays > available) {
        const leaveType = STATIC_LEAVE_TYPES.find(t => t.Id === formState.leaveTypeId);
        errors.submit = `Insufficient ${leaveType?.Title} quota. Available: ${available} days, Requested: ${formState.totalDays} days.`;
      }
    }

    // Check attachment for sick leave > 1 day
    if (formState.leaveTypeId === 3 && formState.totalDays > 1 && !formState.attachment && formState.existingAttachments.length === 0) {
      errors.submit = 'Medical certificate required for sick leave exceeding 1 day';
    }

    return errors;
  };

  handleSubmit = async () => {
    if (!this.state.canEdit) return;

    const errors = this.validateForm();
    if (Object.keys(errors).length > 0) {
      this.setState({ errors });
      return;
    }

    if (window.confirm('Are you sure you want to update this leave request?')) {
      await this.updateLeaveRequest();
    }
  };

  updateLeaveRequest = async () => {
    const { formState, requestId } = this.state;
    
    this.setState({ isSubmitting: true });

    try {
      const updateData: any = {
        LeaveTypeId: formState.leaveTypeId,
        OtherLeaveType: formState.otherLeaveType || null,
        LeaveDurationType: formState.leaveDurationType,
        HalfDayType: formState.halfDayType,
        StartDate: formState.startDate,
        EndDate: formState.endDate,
        TotalDays: formState.totalDays,
        Reason: formState.reason,
        Status: 'Pending on Manager'
      };

      await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(requestId!)
        .update(updateData);

      if (formState.attachment) {
        const buffer = await formState.attachment.arrayBuffer();
        await sp.web.lists
          .getByTitle('Leave Requests')
          .items
          .getById(requestId!)
          .attachmentFiles
          .add(formState.attachment.name, buffer);
      }

      alert('Leave request updated successfully!');
      const url = new URL(window.location.href);
      url.searchParams.delete('RequestID');
      window.location.href = url.toString();
      
    } catch (error: any) {
      this.setState({ errors: { submit: error.message || 'Error updating leave request' } });
    } finally {
      this.setState({ isSubmitting: false });
    }
  };

  deleteAttachment = async (attachmentName: string) => {
    if (!this.state.canEdit) return;
    if (!window.confirm('Are you sure you want to delete this attachment?')) return;

    try {
      await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .getById(this.state.requestId!)
        .attachmentFiles
        .getByName(attachmentName)
        .delete();

      this.setState(prevState => ({
        formState: {
          ...prevState.formState,
          existingAttachments: prevState.formState.existingAttachments.filter(
            (att: any) => att.FileName !== attachmentName
          )
        }
      }));

      alert('Attachment deleted successfully!');
    } catch (error: any) {
      alert('Error deleting attachment: ' + error.message);
    }
  };

  render() {
    const { loading, errorMessage, formState, canEdit, isSubmitting, errors, showHalfDayType, employee, showAttachmentUpload } = this.state;

    if (loading) {
      return (
        <div className={styles.loadingContainer}>
          <div className={styles.spinner}></div>
          <p>Loading leave request...</p>
        </div>
      );
    }

    if (errorMessage && !canEdit) {
      return (
        <div className={styles.errorContainer}>
          <div style={{ textAlign: 'center', padding: '40px' }}>
            <p style={{ color: 'red', fontSize: '16px', marginBottom: '20px' }}>⚠️ {errorMessage}</p>
            <button 
              onClick={() => {
                const url = new URL(window.location.href);
                url.searchParams.delete('RequestID');
                window.location.href = url.toString();
              }}
              style={{ padding: '10px 20px', cursor: 'pointer' }}
            >
              Go Back
            </button>
          </div>
        </div>
      );
    }

    // Only show error summary if there are actual errors
    const hasErrors = Object.keys(errors).length > 0;

    return (
      <div className={styles.applyLeaveFormContent}>
        {!canEdit && errorMessage && (
          <div style={{ backgroundColor: '#fff3cd', color: '#856404', padding: '12px', borderRadius: '4px', marginBottom: '20px', textAlign: 'center' }}>
            ⚠️ {errorMessage}
          </div>
        )}

        {hasErrors && canEdit && (
          <div className={styles.errorSummary}>
            {Object.entries(errors).map(([key, error]) => error && (
              <div key={key} className={styles.errorMessage}>{error}</div>
            ))}
          </div>
        )}

        <div className={styles.applyLeaveForm}>
          {/* Employee Info - Readonly */}
          <div className={styles.applyLeaveRow}>
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Employee Name</label>
              <input type="text" value={employee?.name || ''} disabled className={styles.applyLeaveDisabledField} />
            </div>
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Department</label>
              <input type="text" value={employee?.department || ''} disabled className={styles.applyLeaveDisabledField} />
            </div>
          </div>

          <div className={styles.applyLeaveRow}>
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Country</label>
              <input type="text" value={employee?.country || ''} disabled className={styles.applyLeaveDisabledField} />
            </div>
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Manager</label>
              <input type="text" value={employee?.managerName || 'Not Assigned'} disabled className={styles.applyLeaveDisabledField} />
            </div>
          </div>

          {/* Editable Fields */}
          <div className={styles.applyLeaveRow}>
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Leave Type <span className={styles.applyLeaveRequired}>*</span></label>
              <select 
                value={formState.leaveTypeId} 
                onChange={e => this.handleFieldChange('leaveTypeId', parseInt(e.target.value))} 
                className={styles.applyLeaveSelect}
                disabled={!canEdit}
              >
                <option value="0">Select Leave Type</option>
                {STATIC_LEAVE_TYPES.map(type => (
                  <option key={type.Id} value={type.Id}>{type.Title}</option>
                ))}
              </select>
            </div>

            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Leave Duration <span className={styles.applyLeaveRequired}>*</span></label>
              <select 
                value={formState.leaveDurationType} 
                onChange={e => this.handleFieldChange('leaveDurationType', e.target.value as 'Full Day' | 'Half Day')} 
                className={styles.applyLeaveSelect}
                disabled={!canEdit}
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
                  disabled={!canEdit}
                />
              </div>
            </div>
          )}

          {showHalfDayType && (
            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Half Day Type <span className={styles.applyLeaveRequired}>*</span></label>
                <select 
                  value={formState.halfDayType} 
                  onChange={e => this.handleFieldChange('halfDayType', e.target.value as 'First Half' | 'Second Half')} 
                  className={styles.applyLeaveSelect}
                  disabled={!canEdit}
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
                disabled={!canEdit}
              />
            </div>
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>End Date {formState.leaveDurationType !== 'Half Day' && <span className={styles.applyLeaveRequired}>*</span>}</label>
              <input 
                type="date" 
                value={formState.endDate} 
                disabled={formState.leaveDurationType === 'Half Day' || !canEdit}
                onChange={e => this.handleFieldChange('endDate', e.target.value)} 
                className={styles.applyLeaveInput}
              />
            </div>
          </div>

          <div className={styles.applyLeaveFormGroup}>
            <label className={styles.applyLeaveLabel}>Total Days</label>
            <input 
              type="number" 
              value={formState.totalDays} 
              disabled 
              className={styles.applyLeaveTotalDaysField} 
              step="0.5" 
            />
          </div>

          <div className={styles.applyLeaveFormGroup}>
            <label className={styles.applyLeaveLabel}>Reason for Leave <span className={styles.applyLeaveRequired}>*</span></label>
            <textarea 
              value={formState.reason} 
              onChange={e => this.handleFieldChange('reason', e.target.value)} 
              placeholder="Please provide reason for your leave" 
              rows={4} 
              className={styles.applyLeaveTextarea}
              disabled={!canEdit}
            />
          </div>

          {/* Existing Attachments - Only show if there are attachments */}
          {formState.existingAttachments.length > 0 && (
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Existing Attachments</label>
              <div>
                {formState.existingAttachments.map((att: any, index: number) => (
                  <div key={index} className={styles.applyLeaveFilePreview}>
                    <span>{att.FileName}</span>
                    {canEdit && (
                      <span 
                        className={styles.applyLeaveRemoveFile} 
                        onClick={() => this.deleteAttachment(att.FileName)}
                        style={{ cursor: 'pointer', marginLeft: '10px' }}
                      >
                        ✕
                      </span>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* New Attachment - Only show for Sick Leave > 1 day OR if there are existing attachments */}
          {(showAttachmentUpload || formState.existingAttachments.length > 0) && canEdit && (
            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>
                {formState.leaveTypeId === 3 && formState.totalDays > 1 
                  ? 'Medical Certificate (Required) *' 
                  : 'Add New Attachment (Optional)'}
              </label>
              <div 
                className={styles.applyLeaveFileDropZone} 
                onClick={() => document.getElementById('editFileInput')?.click()}
              >
                {formState.attachment ? (
                  <div className={styles.applyLeaveFilePreview}>
                    <span>{formState.attachment.name}</span>
                    <span 
                      className={styles.applyLeaveRemoveFile} 
                      onClick={(e) => { e.stopPropagation(); this.handleFieldChange('attachment', null); }}
                    >
                      ✕
                    </span>
                  </div>
                ) : (
                  <div className={styles.applyLeaveUploadPrompt}>
                    <span>Click to upload attachment</span>
                    <span className={styles.applyLeaveFileHint}>PDF, JPG, PNG (Max 5MB)</span>
                  </div>
                )}
              </div>
              <input id="editFileInput" type="file" style={{ display: 'none' }} accept=".pdf,.jpg,.jpeg,.png" onChange={this.handleFileChange} />
            </div>
          )}

          {canEdit && (
            <div className={styles.applyLeaveButtonContainer}>
              <button 
                type="button" 
                className={styles.applyLeaveSubmitButton} 
                onClick={this.handleSubmit} 
                disabled={isSubmitting}
              >
                {isSubmitting ? 'Updating...' : 'Update Leave Request'}
              </button>
              <button 
                type="button" 
                onClick={() => {
                  const url = new URL(window.location.href);
                  url.searchParams.delete('RequestID');
                  window.location.href = url.toString();
                }}
                style={{ marginLeft: '12px', padding: '10px 20px', cursor: 'pointer' }}
              >
                Cancel
              </button>
            </div>
          )}
        </div>
      </div>
    );
  }
}