import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './LeaveManagementSystem.module.scss';
import { ILeaveManagementSystemProps } from './ILeaveManagementSystemProps';
import ApplyLeave from './ApplyLeave/ApplyLeave';
import MyDashboard from './MyDashboard/MyDashboard';
import { sp } from '@pnp/sp/presets/all';

interface IEmployeeData {
  Id: number;
  EmployeeName: string;
  Department: string;
  UserImageUrl: string;
}

const LeaveManagementSystem: React.FC<ILeaveManagementSystemProps> = () => {
  const [activeTab, setActiveTab] = useState<string>('myDashboard');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [employee, setEmployee] = useState<IEmployeeData | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  // Check URL for RequestID parameter
  const checkUrlForRequestId = (): boolean => {
    const urlParams = new URLSearchParams(window.location.search);
    const requestIdParam = urlParams.get('RequestID');
    
    if (requestIdParam && requestIdParam.trim() !== '') {
      const requestId = parseInt(requestIdParam, 10);
      if (!isNaN(requestId) && requestId > 0) {
        return true;
      }
    }
    return false;
  };

  // ✅ NEW: Function to handle successful leave submission
  const handleLeaveSubmitted = () => {
    // Dashboard tab par redirect karo
    setActiveTab('myDashboard');
    
    // URL se RequestID parameter hatayein (agar ho to)
    const url = new URL(window.location.href);
    url.searchParams.delete('RequestID');
    window.history.pushState({}, '', url.toString());
  };

  // Fetch employee data for header
  useEffect(() => {
    const fetchEmployeeData = async () => {
      try {
        setIsLoading(true);
        const currentUser = await sp.web.currentUser.get();
        
        const employeeData = await sp.web.lists.getByTitle("Employee Information List").items
          .select(
            "Id", 
            "EmployeeName/Title", 
            "Department/Id",
            "Department/Title", 
            "AttachmentFiles"
          )
          .expand("EmployeeName", "Department", "AttachmentFiles")
          .filter(`EmployeeName/EMail eq '${currentUser.Email}'`)
          .get();

        if (employeeData && employeeData.length > 0) {
          const emp = employeeData[0];
          
          const photoUrl = emp.AttachmentFiles?.length > 0
            ? emp.AttachmentFiles.find((f: any) => f.FileName?.match(/\.(jpg|jpeg|png|gif)$/i))?.ServerRelativeUrl
            : '';

          setEmployee({
            Id: currentUser.Id,
            EmployeeName: emp.EmployeeName?.Title || currentUser.Title,
            Department: emp.Department?.Title || 'Not Assigned',
            UserImageUrl: photoUrl
          });
        } else {
          setEmployee({
            Id: currentUser.Id,
            EmployeeName: currentUser.Title,
            Department: 'Not Assigned',
            UserImageUrl: ''
          });
        }
      } catch (error) {
        console.error('Error fetching employee data:', error);
      } finally {
        setIsLoading(false);
      }
    };

    fetchEmployeeData();

    const hasValidRequestId = checkUrlForRequestId();
    if (hasValidRequestId) {
      setActiveTab('applyLeave');
    }
  }, []);

  const toggleMobileMenu = () => {
    setIsMobileMenuOpen(!isMobileMenuOpen);
  };

  const renderContent = () => {
    switch (activeTab) {
      case 'applyLeave':
        // ✅ CHANGE: Pass the callback function as prop
        return <ApplyLeave onLeaveSubmitted={handleLeaveSubmitted} />;
      case 'myDashboard':
        return <MyDashboard />;
      default:
        return <MyDashboard />;
    }
  };

  return (
    <div className={styles.lmsWrapper}>
      <button className={styles.lmsMobileMenuBtn} onClick={toggleMobileMenu}>
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M3 12H21M3 6H21M3 18H21" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
        </svg>
      </button>

      {isMobileMenuOpen && (
        <div className={styles.lmsOverlay} onClick={() => setIsMobileMenuOpen(false)}></div>
      )}

      <div className={`${styles.lmsSidebar} ${isMobileMenuOpen ? styles.lmsMobileOpen : ''}`}>
        <div className={styles.lmsSidebarHeader}>
          <div className={styles.lmsLogoWrapper}>
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M12 2L2 7L12 12L22 7L12 2Z" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
              <path d="M2 17L12 22L22 17" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
              <path d="M2 12L12 17L22 12" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
          </div>
          <h3>Leave Management</h3>
          <button className={styles.lmsMobileCloseBtn} onClick={toggleMobileMenu}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M18 6L6 18M6 6L18 18" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
            </svg>
          </button>
        </div>

        <nav className={styles.lmsSidebarNav}>
          <button
            className={`${styles.lmsNavItem} ${activeTab === 'myDashboard' ? styles.lmsActive : ''}`}
            onClick={() => {
              setActiveTab('myDashboard');
              setIsMobileMenuOpen(false);
            }}
          >
            <span className={styles.lmsNavIcon}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M3 9L12 3L21 9L12 15L3 9Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M7 12V17L12 20L17 17V12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </span>
            My Dashboard
          </button>

          <button
            className={`${styles.lmsNavItem} ${activeTab === 'applyLeave' ? styles.lmsActive : ''}`}
            onClick={() => {
              setActiveTab('applyLeave');
              setIsMobileMenuOpen(false);
            }}
          >
            <span className={styles.lmsNavIcon}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M12 8V12L15 15M21 12C21 16.9706 16.9706 21 12 21C7.02944 21 3 16.9706 3 12C3 7.02944 7.02944 3 12 3C16.9706 3 21 7.02944 21 12Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
              </svg>
            </span>
            Apply for Leave
          </button>
        </nav>
      </div>

      <div className={styles.lmsMainContent}>
        <div className={styles.lmsHeader}>
          <div className={styles.lmsHeaderContent}>
            <div className={styles.lmsPageTitle}>
              <h2>{activeTab === 'applyLeave' ? 'Apply for Leave' : 'My Dashboard'}</h2>
            </div>
            <div className={styles.lmsUserInfo}>
              {!isLoading && employee && (
                <>
                  <div className={styles.lmsUserAvatar}>
                    {employee.UserImageUrl ? (
                      <img 
                        src={employee.UserImageUrl} 
                        alt={employee.EmployeeName}
                        onError={(e) => {
                          (e.target as HTMLImageElement).src = '';
                          (e.target as HTMLImageElement).style.display = 'none';
                          const parent = (e.target as HTMLImageElement).parentElement;
                          if (parent) {
                            parent.innerHTML = `<svg width="40" height="40" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                              <circle cx="20" cy="20" r="20" fill="#2563eb"/>
                              <path d="M20 10C16.6863 10 14 12.6863 14 16C14 19.3137 16.6863 22 20 22C23.3137 22 26 19.3137 26 16C26 12.6863 23.3137 10 20 10Z" fill="white"/>
                              <path d="M10 30C10 26.6863 12.6863 24 16 24H24C27.3137 24 30 26.6863 30 30V32H10V30Z" fill="white"/>
                            </svg>`;
                          }
                        }}
                      />
                    ) : (
                      <svg width="40" height="40" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="20" cy="20" r="20" fill="#2563eb"/>
                        <path d="M20 10C16.6863 10 14 12.6863 14 16C14 19.3137 16.6863 22 20 22C23.3137 22 26 19.3137 26 16C26 12.6863 23.3137 10 20 10Z" fill="white"/>
                        <path d="M10 30C10 26.6863 12.6863 24 16 24H24C27.3137 24 30 26.6863 30 30V32H10V30Z" fill="white"/>
                      </svg>
                    )}
                  </div>
                  <div className={styles.lmsUserDetails}>
                    <div className={styles.lmsUserName}>{employee.EmployeeName}</div>
                    <div className={styles.lmsUserDept}>{employee.Department}</div>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>

        <div className={styles.lmsContentBody}>
          {renderContent()}
        </div>
      </div>
    </div>
  );
};

export default LeaveManagementSystem;