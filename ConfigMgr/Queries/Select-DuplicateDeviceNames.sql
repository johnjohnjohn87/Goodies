-- WQL query to select devices with the same name
-- Limit to All Systems to get devices without the ConfigMgr client
-- Works best excluding devices named "unknown" (can make that a seperate collection)

select R.ResourceID,R.ResourceType,R.Name,R.SMSUniqueIdentifier,R.ResourceDomainORWorkgroup,R.Client
from SMS_R_System as r
full join SMS_R_System as s1 on s1.ResourceId = r.ResourceId
full join SMS_R_System as s2 on s2.Name = s1.Name
where s1.Name = s2.Name and s1.ResourceId != s2.ResourceId