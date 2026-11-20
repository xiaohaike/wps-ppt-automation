Attribute Module_Name = "Module11"

function ppt模板自动化总函数(){


//-------------------------------------------------主函数------------------------------------------------------
	function ppt模板自动化函数(){
		let datas=Application.InputBox("请用鼠标选择表格区域",undefined, undefined,undefined,undefined,undefined,undefined,8);
		if(!datas){
			alert("函数已退出");
			return
			}
		//文件选择框（选取模板ppt文件）
		let fileDialog = Application.FileDialog(msoFileDialogOpen);
		fileDialog.Title = "请选择需要批量生成的ppt模板";
		let 用户选择结果1=fileDialog.Show();
		
		 if(用户选择结果1==0){
		 	alert("函数已退出");
		 	return
		 }
		 
		 let 用户选择结果2 = fileDialog.SelectedItems.Item(1);
		 if ( !IsPPTFile(用户选择结果2) ){
		 	alert("该文件不是ppt文件，函数已退出");
		 }
		 //如果用户选择完文件，选中用户选择的表格区域，方便后续读取表格中的图片位置来批量插入ppt
		 datas.Select();
		 let wk = CreateObject("kwpp.application");
		 wk.Visible = false;
		 let ppt_app = wk.Presentations.Open(用户选择结果2,0,0,0);//0=false;-1=true
		 let Application_ = Application;
		 
		 let data_obj={
		 	vis_data:processVisibleData(Application_,datas).visibleData,//调用可见区域二维列表函数
		 	ppt_path:用户选择结果2,
		 	select_sheet:Application_.Selection,
		 	sheet_shapes:Application_.ActiveSheet.Shapes,
		 	ActiveSheet:Application_.ActiveSheet,
		 	ppt_obj:ppt_app,
		 	wps:wk,
		 	templateSlide:ppt_app.Slides.Item(1),
		 	photos_sreach_file_array:null	//指定查找照片对象					   
		 }
	
		 ppt_add(data_obj);
		 
	}


//-----------------------------------------------逻辑功能函数----------------------------------------------------
	function ppt_add(data){
		 let ppt=data.ppt_obj;
		 let wps=data.wps;
	
		//判断打开的ppt文件中是否有模板幻灯片
		if(!data.templateSlide ){
			alert("ppt文件中没有模板幻灯片，函数已退出"); 
			return
		}
		if(ppt.Slides.Count>2){
			alert("ppt文件中有多张幻灯片，函数已退出");
			return
			}
	
		//批量生成ppt
		data.vis_data.forEach((undefined,index)=> {
			
		Slide_operate(data,data.vis_data[index]);
			
		});
		
		//删除模板幻灯片
		sleepInWpsMacro(100);
		ppt.Slides.Item(1).Delete();
		
		//删除临时文件夹
		let folderPath = `${Env.GetDesktopPath()}/测试目录`;
		//let fileList = [];
		let fileName = Dir(folderPath + "\\*.*"); // 获取第一个文件
	
		// 循环获取所有文件
		while (fileName != "") {
		if (fileName != "." && fileName != "..") { // 排除特殊目录
			var fullPath = folderPath + "\\" + fileName;
			//遍历删除文件夹中的文件
			FileSystem.Remove(fullPath)
			//fileList.push(fullPath); // 将文件全路径添加到数组
			}
		fileName = Dir(); // 获取下一个文件
		}
		FileSystem.rmdirSync(folderPath);	
		
		
		//将生成的ppt另存至指定文件夹下
		let newFileName = `${Env.GetDesktopPath()}\\批量生成文件__${getCurrentTimeFormatted()}.pptx`;
		ppt.SaveAs(newFileName);
		//不保存当前ppt文件操作并退出
		ppt.Close(false);
		wps.Quit();
		alert(`函数已执行完成，批量文件已自动保存在桌面，文件自动命名为:批量生成文件__${getCurrentTimeFormatted()}.pptx`);
			 
	}

	//修改根据幻灯片模板生成的实例  slide（单页幻灯片） Array_DATA（一维数组：表格中每一行文本内容）
	function Slide_operate(obj,Array_DATA){
		let newSlide=obj.templateSlide.Duplicate();
		/*if (!newSlide) {
		// 处理复制失败的情况
		newSlide=obj.templateSlide.Duplicate();
		}*/
		newSlide.MoveTo(obj.ppt_obj.Slides.Count+1);
		//查找每页幻灯片中的文本框合集
		let textbox_array=getSlideAllTextboxes_Enhanced(newSlide);
		
		//先解析文本占位符
		textbox_array.forEach((box)=>{
			
			let text_box=box.TextFrame.TextRange;
			//正则表达式映射函数:match（返回的匹配组：<n>）,number（返回的匹配组内的匹配字：n）
			let new_text=text_box.Text.replace(/<(\d+)>/g, (match, number) => {
		        // 检查替换映射中是否有对应的替换内容
		        return Array_DATA[number-1] !== undefined ? Array_DATA[number-1] :"";
		    });
			text_box.Text=new_text;
			
			//解析图片占位符
			let match_str1=/@(\d+)@/g;
			if ((match1 =match_str1.exec(text_box.Text)) ) {
				let index=match1[1]-1; //读取搜索符中的索引，并对应表格范围中的二维数组索引
				//清空图片占位符
			    box.TextFrame.TextRange.Text=text_box.Text.replace(/@(\d+)@/g,"");
				if(!Array_DATA[index] || Array_DATA[index]=="" || String(Array_DATA[index]).trim().match(/ID_/)==null ){
					box.Delete();
					return ;
				}
				
			//批量插入图片
			box.Fill.UserPicture(Array_DATA[index]);
	
			}
				
			
			
			//解析图片搜索占位符
			let match_str2=/#(\d+)#/g;
			let match2 =match_str2.exec(text_box.Text);
			//清空图片搜索占位符
			box.TextFrame.TextRange.Text=text_box.Text.replace(/#(\d+)#/g,"");
			let have_match_array=obj.photos_sreach_file_array;
			
	        // 如果没有图片占位符，无需后续操作
			if (match2) {
				
				if(have_match_array=="not choose"){box.Delete();return;}
				let index2=match2[1]-1; //读取搜索符中的索引，并对应表格范围中的二维数组索引
				
			    if(!have_match_array){
			    	
			    	alert("识别到模板文件中有图片占位符，请选择需要批量插入图片的文件夹");
			    	let fd = Application.FileDialog(msoFileDialogFolderPicker);
			        fd.Title = "请选择要读取的文件夹";
			        let 用户点击结果=fd.Show();
			        if(用户点击结果==0){
			        	//如果用户输入了图片搜索符但是取消选择文件夹，则标记跳过
			        	obj.photos_sreach_file_array="not choose";
			        	box.Delete();
			        	return
			        }
			        
			        //获取文件本地路径
			        let folderPath=fd.SelectedItems(1);
			        //obj.photos_file_obj.path=folderPath;
			        //文件夹下子文件路径
			        let fileList = [];
			        let fileName = Dir(folderPath + "\\*.*"); // 获取第一个文件
			    
			        // 循环获取所有文件
			         while (fileName != "") {
				        if (fileName != "." && fileName != "..") { // 排除特殊目录
				            var fullPath = folderPath + "\\" + fileName;
				            fileList.push(fullPath); // 将文件全路径添加到数组
				        }
			        	fileName = Dir(); // 获取下一个文件
			         }
			         
			        obj.photos_sreach_file_array=fileList;
		
			    }
			    
			    //Array_DATA[match]  根据占位符匹配表格内容;  match[1]：搜索出图片占位符的结果数组，索引1为匹配组内的匹配字     
			    let filter_photo_path=findAndRemoveImages(obj.photos_sreach_file_array,Array_DATA[match2[1]-1])[0]  //如果图片占位符匹配到多张照片，只选择第一张
			    
			    if (!filter_photo_path){
			    	box.Delete();
			    	return
			    }
	
			    //批量插入图片
			    box.Fill.UserPicture(filter_photo_path);
			    //如果shape中有边框则隐藏
			    //box.Line.Visible=0;
			    
			}
			
			
		})
	
	}
	
	
	//查找一张幻灯片下所有含有文本的shape对象，并返回查找数组
	function getSlideAllTextboxes_Enhanced(slide) {
	    const textboxList = [];
	    try {
	        if (!slide || !slide.Shapes) {
	            throw new Error("幻灯片对象无效，请确认是WPS PPT的Slide实例");
	        }
	
	        // 核心：递归遍历所有形状（含组合内子形状）
	        function traverseShapes(shapeCollection, parentInfo = "顶层") {
	            for (let i = 1; i <= shapeCollection.Count; i++) {
	                const shape = shapeCollection.Item(i);
	                const shapeType = shape.Type;
	
	                // 1. 若为组合形状，先递归遍历内部子形状
	                if (shapeType === 6) { // 6 = msoGroup（组合形状）
	                    traverseShapes(shape.GroupItems, `组合[${shape.Name}]`);
	                    continue;
	                }
	
	                // 2. 筛选文本框：优先判断HasTextFrame，兼容CanEdit误判
	                if (shape.HasTextFrame) {
	                    const textFrame = shape.TextFrame;
	                    const textRange = textFrame.TextRange;
	
	                    // 兼容处理：即使CanEdit为false，只要有文本内容也纳入（解决占位符误判）
	                    const isEffectiveTextbox = textFrame.CanEdit || (textRange && textRange.Text.trim() !== "");
	                    if (isEffectiveTextbox) {
	                        // 收集文本框信息
	                        textboxList.push(shape);
	                    }
	                }
	            }
	        }
	
	        // 辅助：将形状类型数字转为中文名称（便于理解）
	        function getShapeTypeName(typeNum) {
	            const typeMap = {
	                1: "自动形状",
	                2: "文本框",
	                17: "占位符",
	                20: "图片",
	                6: "组合形状"
	            };
	            return typeMap[typeNum] || `未知类型(${typeNum})`;
	        }
	
	        // 开始遍历顶层形状（自动处理组合内子形状）
	        traverseShapes(slide.Shapes);
	
	        // 结果统计
	        return textboxList;
	
	    } catch (e) {
	        console.error("识别文本框失败：", e.message);
	        alert(`错误：${e.message}`);
	        return [];
	    }
	}
	
	
	/*查找指定文件路径(imageArray：包含文件路径的数组； searchStr：需要查找的文件名)
	!!会修改参数数组
	*/
	function findAndRemoveImages(imageArray, searchStr) {
	    // 定义常见的图片文件扩展名
	    const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.svg'];
	    const foundImages = [];
	    
	    // 倒序遍历数组，避免删除元素后索引混乱
	    for (let i = imageArray.length - 1; i >= 0; i--) {
	        const imagePath = imageArray[i];
	        
	        // 检查是否为字符串类型的路径
	        if (typeof imagePath !== 'string') continue;
	        
	        // 检查是否是图片文件（通过扩展名判断）
	        const isImage = imageExtensions.some(ext => 
	            imagePath.toLowerCase().endsWith(ext)
	        );
	        
	        if (!isImage) continue;
	        
	        // 提取文件名（处理不同操作系统的路径分隔符）
	        const fileName = imagePath.split(/[\\/]/).pop();
	        // 提取不带扩展名的文件名
	        const fileNameWithoutExt = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
	        
	        // 检查文件名是否等于或包含搜索字符串
	        if (fileNameWithoutExt === searchStr || fileNameWithoutExt.includes(searchStr)) {
	            // 将找到的图片添加到结果数组（保持原顺序）
	            foundImages.unshift(imagePath);
	            // 从原数组中删除该元素
	            //imageArray.splice(i, 1);
	        }
	    }
	    
	    return foundImages;
	}
	
	
	
//---------------------------------------------------工具函数---------------------------------------------------

	//返回时间格式化函数，输出时间格式（年-月-日-时-分-秒）
	function getCurrentTimeFormatted(now = new Date()){
		  // 获取时间的各个组成部分
		  const year = now.getFullYear(); // 年 (4位)
		  const month = (now.getMonth() + 1).toString().padStart(2, '0'); // 月 (0-11, 所以+1)
		  const day = now.getDate().toString().padStart(2, '0'); // 日
		  const hours = now.getHours().toString().padStart(2, '0'); // 时
		  const minutes = now.getMinutes().toString().padStart(2, '0'); // 分
		  const seconds = now.getSeconds().toString().padStart(2, '0'); // 秒
		
		  // 拼接成 "年-月-日-时-分-秒" 的格式
		  return `${year}年-${month}月-${day}日-${hours}点-${minutes}分-${seconds}秒`;
	}
	
	
	//将选定区域二维列表转化为可见区域二维列表
	function processVisibleData(ap,Ranges) {
	    if (!Ranges || Ranges.Count === 0) {
	        return { visibleData: [], originalRanges: null };
	    }
	    
	    // 一次性读取所有数据到内存
	    const allData = Ranges.Value2;
	    const rowCount = Ranges.Rows.Count;
	    const colCount = Ranges.Columns.Count;
	    
	    // 记录可见行/列索引（原始表格中的位置）
	    const visibleRowIndices = [];
	    const visibleColIndices = [];
	    
	    //计算浮动图片所在的相对位置
	    let shapes = ap.ActiveSheet.Shapes;
	    let selection = ap.Selection;
	    let startRow = parseInt(selection.Row);  // 选中区域起始行
	    let endRow = startRow + parseInt( selection.Rows.Count ) - 1;  // 选中区域末尾行
	    let startCol = parseInt(ap.Selection.Column);  // 选中区域起始列
	    let endCol = startCol + parseInt( selection.Columns.Count ) - 1;  // 选中区域末尾列
	    
		//先收集区别浮动照片位置并写入内存数组中，同时存入指定临时文件
		let Folder_path=`${Env.GetDesktopPath()}/测试目录`;
		//新建文件夹
		if (!FileSystem.existsSync(Folder_path)) FileSystem.mkdirSync(Folder_path);
	
		for (let i = 1; i <= shapes.Count; i++) {
				let shape=shapes.Item(i);
				let cell = shape.TopLeftCell;
				let row =cell.Row; // 图片所在的绝对行号
				let col =cell.Column; // 图片所在的绝对列号
				let relative_row = row - startRow + 1;
				let relative_col = col - startCol + 1;
				//将包含在选中区域的浮动图片写入二维数组中
				if (shape.Type==13 && row >= startRow && row <= endRow && col >= startCol && col <= endCol){
					let shape_path=`${Folder_path}/${shape.Name}.png`;		
					allData[relative_row-1][relative_col-1] =shape_path;
					shapes.Item(shape.Name).SaveAsPicture(shape_path);
				}
			
			}
				
			
	    // 收集可见行
	    for (let i = 1; i <= rowCount; i++) {
	        if (!Ranges.Rows.Item(i).Hidden) {
	            visibleRowIndices.push(i - 1); // 转换为0基索引
	        }
	    }
	    
	    // 收集可见列
	    for (let j = 1; j <= colCount; j++) {
	        if (!Ranges.Columns.Item(j).Hidden) {
	            visibleColIndices.push(j - 1); // 转换为0基索引
	        }
	    }
	    
	    // 构建可见数据的二维数组
	    const visibleData = visibleRowIndices.map(rowIdx => 
	        visibleColIndices.map(colIdx => allData[rowIdx][colIdx])
	    );
	    
	// visibleData（可见区域）
	    return { visibleData, originalRanges: Ranges };
	}
	
	
	//判断文件是否为ppt
	function IsPPTFile(filePath) {
	    // 常见PPT文件扩展名列表
	    var pptExtensions = ['.ppt', '.pptx', '.pptm', '.pot', '.potx', '.potm', '.pps', '.ppsx', '.ppsm'];
	    
	    // 获取文件扩展名（转换为小写以便比较）
	    var extension = filePath.slice(filePath.lastIndexOf('.')).toLowerCase();
	    
	    // 检查扩展名是否在PPT扩展名列表中
	    return pptExtensions.includes(extension);
	}
	
	
	function sleepInWpsMacro(delayMs) {
		const startTime = Date.now(); // 记录开始时间（毫秒级）
		// 循环等待，直到时间差 >= 目标延迟
		while (Date.now() - startTime < delayMs) {
		    // 空循环：仅等待，不执行任何操作（避免占用过多CPU）
		
		}
		//Console.log(`已延迟 ${delayMs} 毫秒`);
	}
	
//------------------运行主函数------------------------
	ppt模板自动化函数()
}


//----------------加载项按钮函数--------------------
function Workbook_Open() {
	// 遍历所有已存在的功能区，检查是否有"xiaohaike_函数区"
	let bar1 =null;
	let com=Application.CommandBars;
	
    for (let i = 1; i <=com.Count; i++) {
        let currentBar = com(i);
        Console.log(currentBar.Name)
        if (currentBar.Name === "xiaohaike_函数区") {
            bar1 = currentBar; // 找到目标功能区，赋值给bar1
            break; 
        }
    }
    
    if (!bar1) {
        bar1 = Application.CommandBars.Add("xiaohaike_函数区");
    }
    // 确保功能区可见（无论新建还是已存在）
    bar1.Visible = true;
    // 在功能区中添加功能按钮（无论功能区是新建还是已存在）
    let btn1 = bar1.Controls.Add(msoControlButton);
    btn1.Caption = 'xiaohaike_ppt自动化函数'; // 按钮显示名称
    btn1.OnAction = 'ppt模板自动化总函数'; // 绑定宏命令
}





















