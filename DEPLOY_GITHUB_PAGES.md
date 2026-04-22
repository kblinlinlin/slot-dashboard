# GitHub Pages 部署说明

## 重要提醒

这个看板会把 `data/initial-data.js` 中的初始数据一起发布到 GitHub Pages。即使数据已经脱敏，只要 GitHub Pages 地址可访问，其他人也能看到这些数据。

如果这是内部敏感数据，建议使用私有服务器、公司内网、VPN、Cloudflare Access，或在正式版本中增加登录权限。

## 推荐部署方式

1. 在 GitHub 创建一个新仓库，例如：

   ```text
   slot-dashboard
   ```

2. 在本地项目目录添加远程仓库：

   ```powershell
   git remote add origin https://github.com/kblinlinlin/slot-dashboard.git
   ```

3. 推送到 GitHub：

   ```powershell
   git branch -M main
   git push -u origin main
   ```

4. 打开 GitHub 仓库页面：

   ```text
   Settings -> Pages
   ```

5. 在 Build and deployment 中选择：

   ```text
   Source: Deploy from a branch
   Branch: main
   Folder: /root
   ```

6. 保存后等待 GitHub Pages 构建完成。

7. 页面地址通常是：

   ```text
   https://kblinlinlin.github.io/slot-dashboard/
   ```

## 每周上传数据的说明

当前版本支持在网页中上传单周 Excel 数据，并保存到浏览器本地存储。

注意：浏览器本地存储是每个访问者自己的数据，不会自动同步给所有人。如果你需要多人看到同一份更新后的数据，需要后续升级为后端存储，或每周重新生成并提交 `data/initial-data.js`。
