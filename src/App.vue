<template>
  <div class="container">
    <div class="button-group">
      <el-upload
        class="upload-excel"
        action=""
        :auto-upload="false"
        :show-file-list="false"
        accept=".xlsx,.xls"
        @change="handleFileChange"
        ref="uploadRef"
      >
        <template #trigger>
          <el-button type="primary">导入Excel</el-button>
        </template>
      </el-upload>

      <!-- 添加转换按钮 -->
      <el-button
        type="success"
        :disabled="!tableData.length"
        @click="handleTransformData"
      >
        转换数据
      </el-button>

      <!-- 添加导出按钮 -->
      <el-button
        type="warning"
        :disabled="!tableData.length"
        @click="handleExportExcel"
      >
        导出Excel
      </el-button>
    </div>

    <div
      v-if="tableData.length"
      class="table-container"
    >
      <div class="table-info">
        <span>总行数: {{ tableData.length }}</span>
        <span>总列数: {{ tableColumns.length }}</span>
      </div>

      <el-table
        :data="tableData"
        :span-method="objectSpanMethod"
        border
        style="width: 100%"
        height="calc(100vh - 200px)"
        :cell-style="cellStyle"
      >
        <el-table-column
          v-for="(col, index) in tableColumns"
          :key="index"
          :prop="col.prop"
          :label="col.label"
          :min-width="120"
          show-overflow-tooltip
        >
          <template #default="scope">
            <span>{{ scope.row[col.prop] }}</span>
          </template>
        </el-table-column>
      </el-table>
    </div>
  </div>
</template>

<script setup>
  import { ref } from 'vue'
  import { ElMessage } from 'element-plus'
  import * as XLSX from 'xlsx'

  const tableData = ref([])
  const tableColumns = ref([])
  const merges = ref([])
  const spanMethod = ref({})
  const uploadRef = ref(null)
  const originalData = ref(null) // 存储原始数据

  // 添加统计信息的响应式对象
  const statistics = ref({
    show: false,
    actualAttendanceDays: 0,
    lateCount: 0,
    earlyLeaveCount: 0,
    lateTimes: [],
    earlyLeaveTimes: [],
    totalLateMinutes: 0,
    totalEarlyLeaveMinutes: 0
  })

  // 处理文件上传
  const handleFileChange = (file) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        const worksheet = workbook.Sheets[workbook.SheetNames[0]]

        // 获取合并单元格信息
        merges.value = worksheet['!merges'] || []

        // 转换数据，保留所有行
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: '', // 设置空单元格的默认值
          raw: false  // 保持原始格式
        })

        // 保存原始数据
        originalData.value = jsonData

        // 处理并显示数据
        processAndDisplayData(jsonData)

        ElMessage.success('文件导入成功')
      } catch (error) {
        ElMessage.error('文件解析失败')
        console.error(error)
      }
      // 清除上传文件
      uploadRef.value.clearFiles()
    }
    reader.readAsArrayBuffer(file.raw)
  }

  // 计算两个时间之间的小时差
  const calculateHoursDifference = (time1, time2) => {
    try {
      // 解析时间字符串（假设格式为 "HH:mm" 或 "HH:mm:ss"）
      const [hours1, minutes1] = time1.split(':').map(Number)
      const [hours2, minutes2] = time2.split(':').map(Number)

      // 计算分钟差
      const totalMinutes1 = hours1 * 60 + minutes1
      const totalMinutes2 = hours2 * 60 + minutes2

      // 计算小时差
      return Math.abs(totalMinutes2 - totalMinutes1) / 60
    } catch (error) {
      return 0
    }
  }

  // 转换数据处理函数
  const handleTransformData = () => {
    if (!originalData.value || !originalData.value.length) {
      ElMessage.warning('没有可转换的数据')
      return
    }

    try {
      // 去掉前三行数据
      let transformedData = originalData.value.slice(3)

      // 需要保留的列索引（注意：索引从0开始）
      const keepColumns = [0, 1, 7, 8, 9, 10, 11, 12, 13, 14]  // 对应第1,2,8,9,10,12,13,14,15列

      // 筛选列
      transformedData = transformedData.map(row => {
        return keepColumns.map(colIndex => row[colIndex] || '')
      })

      // 修改第一行的列标题
      if (transformedData.length > 0) {
        transformedData[0][0] = '日期'  // 第一列改为"日期"
        transformedData[0][1] = '姓名'  // 第二列改为"姓名"
      }

      // 保存第一行（列标题）
      const headerRow = transformedData[0]

      // 反转数据（除了第一行）
      const dataRows = transformedData.slice(1)
      const reversedData = dataRows.reverse()

      // 重新组合数据：标题行 + 反转后的数据
      transformedData = [headerRow, ...reversedData]

      // 更新合并单元格信息并调整行号
      updateMergesAfterTransform(keepColumns, dataRows.length)

      // 计算统计信息
      const stats = {
        actualAttendanceDays: 0,
        lateCount: 0,
        earlyLeaveCount: 0,
        lateTimes: [],
        earlyLeaveTimes: [],
        totalLateMinutes: 0,
        totalEarlyLeaveMinutes: 0
      }

      transformedData.forEach((row, index) => {
        if (index === 0) return // 跳过标题行

        const startTime = row[3] // 第四列（上班时间）
        const endTime = row[4]   // 第五列（下班时间）
        const status = row[9]    // 第十列
        const date = row[0]      // 第一列（日期）

        // 计算出勤天数
        if (startTime && endTime) {
          const hoursDiff = calculateHoursDifference(startTime, endTime)
          if (hoursDiff >= 7) {
            stats.actualAttendanceDays += 1
          } else if (hoursDiff >= 3) {
            stats.actualAttendanceDays += 0.5
          }
        }

        if (status) {
          // 处理迟到
          if (status.includes('迟到')) {
            stats.lateCount++

            // 提取迟到分钟数
            let lateMinutes = 0
            const lateMatches = status.match(/迟到.*?(\d+)分钟/g) || []

            lateMatches.forEach(match => {
              const minutes = match.match(/\d+/g) || []
              minutes.forEach(min => {
                lateMinutes += parseInt(min, 10)
              })
            })

            stats.totalLateMinutes += lateMinutes
            stats.lateTimes.push(`${date}(${lateMinutes}分钟)`)
          }

          // 处理早退
          if (status.includes('早退')) {
            stats.earlyLeaveCount++

            // 提取早退分钟数
            let earlyMinutes = 0
            const earlyMatches = status.match(/早退.*?(\d+)分钟/g) || []

            earlyMatches.forEach(match => {
              const minutes = match.match(/\d+/g) || []
              minutes.forEach(min => {
                earlyMinutes += parseInt(min, 10)
              })
            })

            stats.totalEarlyLeaveMinutes += earlyMinutes
            stats.earlyLeaveTimes.push(`${date}(${earlyMinutes}分钟)`)
          }
        }
      })

      // 添加统计行到数据末尾
      const summaryRow = Array(transformedData[0].length).fill('')
      summaryRow[0] = '统计信息'
      summaryRow[1] = `实际出勤: ${stats.actualAttendanceDays.toFixed(1)}天`
      summaryRow[2] = `迟到: ${stats.lateCount}次(共${stats.totalLateMinutes}分钟)`
      summaryRow[3] = `早退: ${stats.earlyLeaveCount}次(共${stats.totalEarlyLeaveMinutes}分钟)`
      transformedData.push(summaryRow)

      // 更新统计信息（移除所有统计项，因为已经显示在表格中）
      statistics.value = {
        show: false // 不再显示统计信息面板
      }

      // 重新处理并显示数据
      processAndDisplayData(transformedData)

      ElMessage.success('数据转换成功')
    } catch (error) {
      ElMessage.error('数据转换失败')
      console.error(error)
    }
  }

  // 更新合并单元格信息
  const updateMergesAfterTransform = (keepColumns, dataLength) => {
    if (!merges.value.length) return

    // 创建列映射关系（原始列索引 -> 新列索引）
    const columnMapping = {}
    keepColumns.forEach((oldIndex, newIndex) => {
      columnMapping[oldIndex] = newIndex
    })

    // 过滤并更新合并单元格信息
    merges.value = merges.value
      .filter(merge => {
        // 检查合并单元格是否在保留的列中
        return keepColumns.includes(merge.s.c) && keepColumns.includes(merge.e.c)
      })
      .map(merge => {
        // 更新列索引
        const newMerge = {
          s: { ...merge.s },
          e: { ...merge.e }
        }

        // 调整行号（去掉前三行）
        newMerge.s.r -= 3
        newMerge.e.r -= 3

        // 反转行号（除了第一行）
        if (newMerge.s.r > 0) {
          newMerge.s.r = dataLength - newMerge.s.r
        }
        if (newMerge.e.r > 0) {
          newMerge.e.r = dataLength - newMerge.e.r
        }

        // 确保起始行号小于结束行号
        if (newMerge.s.r > newMerge.e.r) {
          [newMerge.s.r, newMerge.e.r] = [newMerge.e.r, newMerge.s.r]
        }

        // 更新列号
        newMerge.s.c = columnMapping[merge.s.c]
        newMerge.e.c = columnMapping[merge.e.c]

        return newMerge
      })
      .filter(merge => merge.s.r >= 0) // 移除无效的合并单元格
  }

  // 处理并显示数据
  const processAndDisplayData = (jsonData) => {
    // 设置列
    if (jsonData.length > 0) {
      const maxCols = jsonData[0].length
      tableColumns.value = Array(maxCols).fill(0).map((_, index) => ({
        prop: `col${index}`,
        label: `Column ${index + 1}` // 由于去掉了表头，使用默认列名
      }))
    }

    // 设置数据
    tableData.value = jsonData.map(row => {
      const obj = {}
      tableColumns.value.forEach((col, index) => {
        obj[col.prop] = row[index] || ''
      })
      return obj
    })

    // 处理合并单元格
    processMerges()
  }

  // 处理合并单元格
  const processMerges = () => {
    const spanMethodObj = {}

    merges.value.forEach(merge => {
      const { s, e } = merge
      const rowspan = e.r - s.r + 1
      const colspan = e.c - s.c + 1

      for (let row = s.r; row <= e.r; row++) {
        for (let col = s.c; col <= e.c; col++) {
          if (row === s.r && col === s.c) {
            spanMethodObj[`${row}-${col}`] = { rowspan, colspan }
          } else {
            spanMethodObj[`${row}-${col}`] = { rowspan: 0, colspan: 0 }
          }
        }
      }
    })

    spanMethod.value = spanMethodObj
  }

  // 合并单元格方法
  const objectSpanMethod = ({ row, column, rowIndex, columnIndex }) => {
    const key = `${rowIndex}-${columnIndex}`
    if (spanMethod.value[key]) {
      return spanMethod.value[key]
    }
    return {
      rowspan: 1,
      colspan: 1
    }
  }

  // 单元格样式
  const cellStyle = ({ row, column, rowIndex }) => {
    const baseStyle = {
      textAlign: 'center'
    }

    // 处理最后一行（统计行）的样式
    const lastIndex = tableData.value.length - 1
    if (rowIndex === lastIndex) {
      return {
        ...baseStyle,
        backgroundColor: '#e6f1fc',  // 浅蓝色背景
        color: '#409EFF',           // 蓝色文字
        fontWeight: 'bold',         // 加粗
        fontSize: '14px'            // 稍大字号
      }
    }

    // 处理合并单元格样式
    const key = `${rowIndex}-${column.columnIndex}`
    if (spanMethod.value[key] && (spanMethod.value[key].rowspan > 1 || spanMethod.value[key].colspan > 1)) {
      baseStyle.backgroundColor = '#f5f7fa'
    }

    // 处理迟到早退标红
    if (column.columnIndex === 9) {
      const value = row[`col${column.columnIndex}`]
      if (value && (value.includes('迟到') || value.includes('早退'))) {
        baseStyle.color = '#ff4444'
        baseStyle.fontWeight = 'bold'
      }
    }

    return baseStyle
  }

  // 生成文件名
  const generateFileName = (data) => {
    try {
      const yearMonth = data[1]['col0'] || '' // 第一列第二行（日期）
      const name = data[2]['col1'] || ''      // 第二列第三行（姓名）

      // 从日期中提取年月
      const dateMatch = yearMonth.match(/(\d{4}).*?(\d{1,2})/)
      if (dateMatch) {
        const [, year, month] = dateMatch
        const paddedMonth = month.padStart(2, '0')
        return `${year}-${paddedMonth}-${name}`
      }

      // 如果无法提取日期，使用默认文件名
      return `attendance-${new Date().getTime()}`
    } catch (error) {
      return `attendance-${new Date().getTime()}`
    }
  }

  // 导出Excel处理函数
  const handleExportExcel = () => {
    try {
      // 准备导出数据（保留所有数据行）
      const exportData = tableData.value.map(row => {
        const rowData = {}
        tableColumns.value.forEach((col, index) => {
          rowData[`col${index}`] = row[col.prop] // 使用简单的列标识符
        })
        return rowData
      })

      // 创建工作簿
      const wb = XLSX.utils.book_new()

      // 转换数据为工作表（不使用表头）
      const ws = XLSX.utils.json_to_sheet(exportData, {
        header: tableColumns.value.map((_, index) => `col${index}`),
        skipHeader: true // 跳过表头
      })

      // 设置列宽
      const colWidths = tableColumns.value.map(() => ({ wch: 15 }))
      ws['!cols'] = colWidths

      // 添加合并单元格信息（保持原始行号）
      ws['!merges'] = merges.value.map(merge => ({
        s: { r: merge.s.r, c: merge.s.c },
        e: { r: merge.e.r, c: merge.e.c }
      }))

      // 设置单元格样式和边框
      const lastRowIndex = exportData.length - 1
      for (let row = 0; row < exportData.length; row++) {
        for (let col = 0; col < tableColumns.value.length; col++) {
          const cellRef = XLSX.utils.encode_cell({ r: row, c: col })
          if (!ws[cellRef]) continue

          // 基础样式
          ws[cellRef].s = {
            alignment: { horizontal: 'center', vertical: 'center' },
            font: { name: 'Arial', sz: 11 },
            border: {
              top: { style: 'thin', color: { rgb: '000000' } },
              bottom: { style: 'thin', color: { rgb: '000000' } },
              left: { style: 'thin', color: { rgb: '000000' } },
              right: { style: 'thin', color: { rgb: '000000' } }
            }
          }

          // 第一行样式（浅绿色背景）
          if (row === 0) {
            ws[cellRef].s.fill = {
              fgColor: { rgb: 'E6EFDC' }, // 浅绿色背景
              patternType: 'solid'
            }
            ws[cellRef].s.font.bold = true
          }

          // 最后一行样式（统计行，浅蓝色背景）
          if (row === lastRowIndex) {
            ws[cellRef].s.fill = {
              fgColor: { rgb: 'E6F1FC' },
              patternType: 'solid'
            }
            ws[cellRef].s.font.bold = true
            ws[cellRef].s.font.color = { rgb: '409EFF' }
          }

          // 迟到早退红色字体
          if (col === 9) {
            const value = exportData[row][`col${col}`]
            if (value && (value.includes('迟到') || value.includes('早退'))) {
              ws[cellRef].s.font.color = { rgb: 'FF4444' }
              ws[cellRef].s.font.bold = true
            }
          }
        }
      }

      // 生成文件名
      const fileName = generateFileName(tableData.value)

      // 将工作表添加到工作簿
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')

      // 导出文件
      XLSX.writeFile(wb, `${fileName}.xlsx`)

      ElMessage.success('导出成功')
    } catch (error) {
      console.error('导出失败:', error)
      ElMessage.error('导出失败')
    }
  }
</script>

<style scoped>
  .container {
    padding: 20px;
    max-width: 1400px;
    margin: 0 auto;
    height: 100vh;
    display: flex;
    flex-direction: column;
  }

  .button-group {
    display: flex;
    gap: 16px;
    margin-bottom: 20px;
    align-items: center;
  }

  .upload-excel {
    display: inline-block;
  }

  .table-container {
    flex: 1;
    display: flex;
    flex-direction: column;
    min-height: 0;
  }

  .table-info {
    margin-bottom: 10px;
    display: flex;
    gap: 20px;
    font-size: 14px;
    color: #606266;
  }

  :deep(.el-table__body-wrapper) {
    overflow-y: auto;
  }

  :deep(.el-table__header-wrapper) {
    background-color: #f5f7fa;
  }

  :deep(.el-table th) {
    background-color: #f5f7fa;
    font-weight: bold;
  }

  :deep(.el-table td),
  :deep(.el-table th) {
    padding: 8px 0;
  }
</style> 